/**
 * Column Mapper - Ada's intelligent column mapping service
 * VERSION: 3.3.0 - target_key REMOVED
 * 
 * ARCHITECTURE (Dec 2025 - CANONICAL):
 * - pf_column_name is the ONLY semantic key in database
 * - target is the ONLY canonical field in API response
 * - ada_payroll_column_dictionary = Source of truth for AMOUNTS
 * - ada_payroll_dimensions = Source of truth for DIMENSIONS
 * 
 * RESPONSE SHAPE:
 * {
 *   raw_header: string,        // what the uploaded file calls it
 *   kind: "amount" | "dimension" | "ambiguous" | null,
 *   target: string | null,     // PF canonical name (pf_column_name)
 *   source: "saved" | "amount" | "dimension" | "fuzzy" | "unmapped" | "ambiguous",
 *   confidence: number,
 *   gl_account?: string,       // amounts only
 *   gl_account_name?: string   // amounts only
 * }
 * 
 * LOOKUP ORDER:
 * 1. ada_customer_column_mappings.pf_column_name (saved) → target
 * 2. ada_payroll_column_dictionary.pf_column_name (amounts) → target
 * 3. ada_payroll_dimensions.normalized_dimension (dimensions) → target
 * 4. Fuzzy match → target
 * 5. Otherwise → target = null, source = "unmapped"
 */

const VERSION = "3.3.0";

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") || "https://jgciqwzwacaesqjaoadc.supabase.co";
const SUPABASE_SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

// NOTE: BLOCKED_TARGETS and BLOCKED_HEADERS have been removed.
// Column filtering is now database-driven via include_in_matrix field
// in ada_customer_column_mappings table. Columns with include_in_matrix=false
// will be excluded from PR_Data_Clean by the client.

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

// ============================================================================
// Types
// ============================================================================

interface ColumnMapperRequest {
  headers?: string[];
  crm_company_id?: string | null;
  module?: string;
  mappings?: Array<{ raw_header: string; target: string; kind?: MappingKind }>;
  action: "analyze" | "save" | "get_options" | "get_expense_taxonomy" | "bootstrap" | "debug" | "get_dimensions" | "validate_email";
  installation_key?: string;
  provider?: string;
  email?: string;
}

type MappingKind = "amount" | "dimension" | "ambiguous";
type MappingSource = "saved" | "amount" | "dimension" | "fuzzy" | "unmapped" | "ambiguous";

interface MappingResult {
  raw_header: string;
  kind: MappingKind | null;
  target: string | null;        // PF canonical name (pf_column_name or normalized_dimension)
  source: MappingSource;
  confidence: number;
  gl_account?: string | null;
  gl_account_name?: string | null;
  include_in_matrix?: boolean;      // Whether to include in PR_Data_Clean
  expense_bucket?: string | null;   // FIXED, VARIABLE, TAXES, BENEFITS, OTHER
  // For ambiguous results, include both options
  amount_option?: { target: string; confidence: number } | null;
  dimension_option?: { target: string; confidence: number } | null;
}

interface DictionaryMatch {
  target: string;   // pf_column_name
  confidence: number;
  isExact: boolean;
}

interface SaveMapping {
  raw_header: string;
  target: string;   // CANONICAL: pf_column_name
  kind?: MappingKind;
}

// ============================================================================
// Helpers
// ============================================================================

function getSupabaseClient() {
  if (!SUPABASE_SERVICE_KEY) {
    throw new Error("SUPABASE_SERVICE_ROLE_KEY not configured");
  }
  return createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
}

/**
 * Normalize a header string for matching
 */
function normalizeHeader(header: string): string {
  return (header || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ");
}

/**
 * Calculate similarity score between two strings (0-1)
 */
function similarity(a: string, b: string): number {
  const aLen = a.length;
  const bLen = b.length;
  if (aLen === 0 || bLen === 0) return 0;
  
  // Check for containment
  if (a.includes(b) || b.includes(a)) {
    return Math.min(aLen, bLen) / Math.max(aLen, bLen);
  }
  
  // Simple word overlap
  const aWords = new Set(a.split(" "));
  const bWords = new Set(b.split(" "));
  let overlap = 0;
  for (const word of aWords) {
    if (bWords.has(word)) overlap++;
  }
  
  return overlap / Math.max(aWords.size, bWords.size);
}

/**
 * Apply dimension-specific overrides for known patterns
 * Example: If both "Department" and "Department Description" exist,
 *          map them to Department_Code and Department_Name respectively
 */
function applyDimensionOverrides(mappings: MappingResult[]): MappingResult[] {
  const byHeader = new Map<string, MappingResult>();
  for (const m of mappings) byHeader.set(m.raw_header, m);
  
  const hasDept = byHeader.has("Department");
  const hasDeptDesc = byHeader.has("Department Description");
  
  // If "Department Description" exists, it should win Department_Name,
  // and "Department" should become Department_Code.
  if (hasDept && hasDeptDesc) {
    const dept = byHeader.get("Department")!;
    const deptDesc = byHeader.get("Department Description")!;
    
    // Force Department Description -> Department_Name
    deptDesc.kind = "dimension";
    deptDesc.target = "Department_Name";
    deptDesc.source = "override";
    deptDesc.confidence = 1.0;
    
    // Force Department -> Department_Code
    dept.kind = "dimension";
    dept.target = "Department_Code";
    dept.source = "override";
    dept.confidence = 1.0;
    
    console.log("[Override] Applied Department/Department Description override");
  }
  
  return mappings;
}

// ============================================================================
// Lookup Functions
// ============================================================================

/**
 * Normalize mapping_type to MappingKind
 * Default to "amount" if unknown
 */
function normalizeKind(input: unknown): MappingKind {
  if (input === "dimension") return "dimension";
  return "amount"; // default fallback
}

/**
 * Look up saved mappings for a company and module
 * Priority 1: Company-specific learned mappings
 * 
 * CANONICAL: Uses pf_column_name as the ONLY semantic key.
 * Reads mapping_type to determine kind (amount vs dimension).
 */
async function getSavedMappings(
  supabase: ReturnType<typeof getSupabaseClient>,
  companyId: string | null,
  module: string,
  headers: string[]
): Promise<Map<string, MappingResult>> {
  const results = new Map<string, MappingResult>();
  
  if (!companyId) {
    console.log("[Saved] No crm_company_id provided, skipping saved mappings lookup");
    return results;
  }
  
  if (!headers || headers.length === 0) {
    console.log("[Saved] No headers provided");
    return results;
  }
  
  try {
    console.log(`[Saved] Looking up saved mappings for company ${companyId}, module ${module}`);
    
    // Build lookup maps
    interface SavedRow {
      raw_header: string;
      pf_column_name: string | null;
      mapping_type: string | null;
      confidence: number | null;
      include_in_matrix: boolean | null;
      expense_bucket: string | null;
    }
    const exactLookup = new Map<string, SavedRow>();
    const normalizedLookup = new Map<string, SavedRow>();
    
    // Select pf_column_name and mapping_type for kind determination
    // Also include include_in_matrix and expense_bucket for filtering (if columns exist)
    let { data, error } = await supabase
      .from("ada_customer_column_mappings")
      .select("raw_header, pf_column_name, mapping_type, confidence, include_in_matrix, expense_bucket")
      .eq("crm_company_id", companyId)
      .eq("module", module);
    
    if (error) {
      console.error("[Saved] Database error:", error);
      console.error("[Saved] Error details:", JSON.stringify(error));
      // If error is about missing columns, try without the new columns
      if (error.message?.includes("column") || error.code === "42703") {
        console.log("[Saved] Retrying without include_in_matrix and expense_bucket columns...");
        const fallbackResult = await supabase
          .from("ada_customer_column_mappings")
          .select("raw_header, pf_column_name, mapping_type, confidence")
          .eq("crm_company_id", companyId)
          .eq("module", module);
        
        if (fallbackResult.error) {
          console.error("[Saved] Fallback query also failed:", fallbackResult.error);
          return results;
        }
        
        // Use fallback data with defaults for new columns
        data = (fallbackResult.data || []).map(row => ({
          ...row,
          include_in_matrix: true,  // Default
          expense_bucket: null
        }));
        error = null;
      } else {
        return results;
      }
    }
    
    if (!data || data.length === 0) {
      console.log(`[Saved] No saved mappings found for company ${companyId}, module ${module}`);
      return results;
    }
    
    console.log(`[Saved] Found ${data.length} rows in ada_customer_column_mappings`);
    
    for (const row of data) {
      if (!row.raw_header || !row.pf_column_name) {
        console.log(`[Saved] Skipping row with missing pf_column_name:`, row);
        continue;
      }
      
      exactLookup.set(row.raw_header, row);
      normalizedLookup.set(row.raw_header.toLowerCase().trim(), row);
    }
    
    console.log(`[Saved] Built lookups with ${exactLookup.size} entries`);
    
    // Match incoming headers
    for (const header of headers) {
      let savedRow = exactLookup.get(header);
      
      if (!savedRow) {
        savedRow = normalizedLookup.get(header.toLowerCase().trim());
      }
      
      if (savedRow && savedRow.pf_column_name) {
        const kind = normalizeKind(savedRow.mapping_type);
        results.set(header, {
          raw_header: header,
          kind: kind,
          target: savedRow.pf_column_name,
          source: "saved",
          confidence: Number(savedRow.confidence ?? 1),
          include_in_matrix: savedRow.include_in_matrix ?? true,  // Default true for backwards compatibility
          expense_bucket: savedRow.expense_bucket || null
        });
        console.log(`[Saved] Matched "${header}" → "${savedRow.pf_column_name}" (${kind}, include=${savedRow.include_in_matrix ?? true}, bucket=${savedRow.expense_bucket || 'null'})`);
      }
    }
    
    console.log(`[Saved] Returning ${results.size} matched saved mappings`);
    return results;
  } catch (e) {
    console.error("[Saved] Exception:", e);
    return results;
  }
}

/**
 * Look up AMOUNT matches from ada_payroll_column_dictionary
 * Returns best match (exact or fuzzy) for each header
 * 
 * IMPORTANT: Uses pf_column_name as the target (not normalized_key)
 */
async function getAmountMatches(
  supabase: ReturnType<typeof getSupabaseClient>,
  module: string,
  headers: string[]
): Promise<Map<string, DictionaryMatch>> {
  const results = new Map<string, DictionaryMatch>();
  
  try {
    const { data: dictionary, error } = await supabase
      .from("ada_payroll_column_dictionary")
      .select("data_source_name, pf_column_name")
      .eq("module", module);
    
    if (error) {
      console.error("Error fetching amount dictionary:", error);
      return results;
    }
    
    if (!dictionary || dictionary.length === 0) {
      console.log(`[Amount] No dictionary entries for module: ${module}`);
      return results;
    }
    
    console.log(`[Amount] Loaded ${dictionary.length} entries for module: ${module}`);
    
    // Build lookup: data_source_name → pf_column_name
    const dictLookup = new Map<string, string>();
    for (const entry of dictionary) {
      if (!entry.data_source_name || !entry.pf_column_name) continue;
      const normalized = normalizeHeader(entry.data_source_name);
      dictLookup.set(normalized, entry.pf_column_name);
    }
    
    // Match headers
    for (const header of headers) {
      const normalizedHeader = normalizeHeader(header);
      
      // Exact match
      const exactMatch = dictLookup.get(normalizedHeader);
      if (exactMatch) {
        results.set(header, {
          target: exactMatch,  // pf_column_name
          confidence: 0.95,
          isExact: true
        });
        continue;
      }
      
      // Fuzzy match
      let bestMatch: { target: string; score: number } | null = null;
      for (const [term, pfColumnName] of dictLookup) {
        const score = similarity(normalizedHeader, term);
        if (score > 0.5 && (!bestMatch || score > bestMatch.score)) {
          bestMatch = { target: pfColumnName, score };
        }
      }
      
      if (bestMatch) {
        results.set(header, {
          target: bestMatch.target,  // pf_column_name
          confidence: bestMatch.score * 0.8,
          isExact: false
        });
      }
    }
    
    console.log(`[Amount] Found ${results.size} matches`);
    return results;
  } catch (e) {
    console.error("getAmountMatches error:", e);
    return results;
  }
}

/**
 * Look up DIMENSION matches from ada_payroll_dimensions
 * Returns best match (exact or fuzzy) for each header
 * 
 * Uses normalized_dimension as the target
 */
async function getDimensionMatches(
  supabase: ReturnType<typeof getSupabaseClient>,
  headers: string[]
): Promise<Map<string, DictionaryMatch>> {
  const results = new Map<string, DictionaryMatch>();
  
  try {
    const { data: dimensions, error } = await supabase
      .from("ada_payroll_dimensions")
      .select("raw_term, normalized_dimension");
    
    if (error) {
      console.error("Error fetching dimensions:", error);
      return results;
    }
    
    if (!dimensions || dimensions.length === 0) {
      console.log("[Dimension] No dimension entries found");
      return results;
    }
    
    console.log(`[Dimension] Loaded ${dimensions.length} dimension entries`);
    
    // Build lookup: raw_term → normalized_dimension
    const dimLookup = new Map<string, string>();
    for (const dim of dimensions) {
      if (!dim.raw_term || !dim.normalized_dimension) continue;
      const normalized = normalizeHeader(dim.raw_term);
      dimLookup.set(normalized, dim.normalized_dimension);
    }
    
    // Match headers
    for (const header of headers) {
      const normalizedHeader = normalizeHeader(header);
      
      // Exact match
      const exactMatch = dimLookup.get(normalizedHeader);
      if (exactMatch) {
        results.set(header, {
          target: exactMatch,  // normalized_dimension
          confidence: 0.95,
          isExact: true
        });
        continue;
      }
      
      // Fuzzy match
      let bestMatch: { target: string; score: number } | null = null;
      for (const [term, normalizedDimension] of dimLookup) {
        const score = similarity(normalizedHeader, term);
        if (score > 0.5 && (!bestMatch || score > bestMatch.score)) {
          bestMatch = { target: normalizedDimension, score };
        }
      }
      
      if (bestMatch) {
        results.set(header, {
          target: bestMatch.target,  // normalized_dimension
          confidence: bestMatch.score * 0.7,
          isExact: false
        });
      }
    }
    
    console.log(`[Dimension] Found ${results.size} matches`);
    return results;
  } catch (e) {
    console.error("getDimensionMatches error:", e);
    return results;
  }
}

/**
 * Look up GL accounts for targets from ada_customer_gl_mappings
 * 
 * CANONICAL: Uses pf_column_name as the ONLY join key.
 * normalized_key column has been DROPPED.
 */
async function getGLMappings(
  supabase: ReturnType<typeof getSupabaseClient>,
  companyId: string,
  module: string,
  targets: string[]
): Promise<Map<string, { gl_account: string; gl_account_name?: string }>> {
  const results = new Map<string, { gl_account: string; gl_account_name?: string }>();
  
  if (!companyId || targets.length === 0) return results;
  
  try {
    // CANONICAL: Only use pf_column_name (normalized_key is DROPPED)
    const { data, error } = await supabase
      .from("ada_customer_gl_mappings")
      .select("pf_column_name, gl_account, gl_account_name, priority")
      .eq("crm_company_id", companyId)
      .eq("module", module)
      .in("pf_column_name", targets)
      .order("priority", { ascending: false });
    
    if (error) {
      console.error("[GL] Error fetching GL mappings:", error);
      return results;
    }
    
    if (!data || data.length === 0) {
      console.log(`[GL] No GL mappings found for company ${companyId}`);
      return results;
    }
    
    // Take highest priority match for each pf_column_name
    const seen = new Set<string>();
    
    for (const row of data) {
      if (!row.pf_column_name || seen.has(row.pf_column_name)) continue;
      
      results.set(row.pf_column_name, {
        gl_account: row.gl_account,
        gl_account_name: row.gl_account_name || undefined
      });
      seen.add(row.pf_column_name);
    }
    
    console.log(`[GL] Found ${results.size} GL mappings`);
    return results;
  } catch (e) {
    console.error("[GL] Exception:", e);
    return results;
  }
}

/**
 * Save confirmed mappings for a company and module
 * 
 * CANONICAL: Saves target as pf_column_name (the ONLY semantic key).
 * normalized_key column has been DROPPED.
 */
async function saveMappings(
  supabase: ReturnType<typeof getSupabaseClient>,
  companyId: string,
  module: string,
  mappings: SaveMapping[]
): Promise<{ success: boolean; saved: number; errors: string[]; rejected: string[] }> {
  const errors: string[] = [];
  const rejected: string[] = [];
  let saved = 0;
  
  // VALIDATION: Fetch valid pf_column_names from dictionary
  const { data: validAmounts } = await supabase
    .from("ada_payroll_column_dictionary")
    .select("pf_column_name")
    .eq("module", module);
  
  const { data: validDimensions } = await supabase
    .from("ada_payroll_dimensions")
    .select("normalized_dimension");
  
  const validTargets = new Set([
    ...(validAmounts || []).map((r: any) => r.pf_column_name),
    ...(validDimensions || []).map((r: any) => r.normalized_dimension)
  ]);
  
  console.log(`[Save] Validating against ${validTargets.size} valid targets`);
  
  for (const mapping of mappings) {
    if (!mapping.raw_header || !mapping.target) continue;
    
    // NOTE: BLOCKED_TARGETS validation removed - now handled by include_in_matrix in database
    
    // VALIDATION: Reject invalid targets
    if (!validTargets.has(mapping.target)) {
      rejected.push(`"${mapping.raw_header}" → "${mapping.target}" (invalid target)`);
      console.warn(`[Save] REJECTED invalid target: ${mapping.target} for raw_header "${mapping.raw_header}"`);
      continue;
    }
    
    try {
      const { error } = await supabase
        .from("ada_customer_column_mappings")
        .upsert({
          crm_company_id: companyId,
          module: module,
          raw_header: mapping.raw_header,
          pf_column_name: mapping.target,  // CANONICAL: only semantic key
          mapping_type: mapping.kind || "amount",
          confidence: 1.0,
          source: "ada_confirmed",
          updated_at: new Date().toISOString()
        }, {
          onConflict: "crm_company_id,module,raw_header"
        });
      
      if (error) {
        errors.push(`Failed to save ${mapping.raw_header}: ${error.message}`);
      } else {
        saved++;
        console.log(`[Save] Saved: ${mapping.raw_header} → ${mapping.target}`);
      }
    } catch (e) {
      errors.push(`Error saving ${mapping.raw_header}: ${e}`);
    }
  }
  
  console.log(`[Save] Saved ${saved} mappings, ${errors.length} errors, ${rejected.length} rejected`);
  return { 
    success: errors.length === 0 && rejected.length === 0, 
    saved, 
    errors,
    rejected 
  };
}

// ============================================================================
// Main Handler
// ============================================================================

serve(async (req) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }
  
  try {
    const body: ColumnMapperRequest = await req.json();
    const { headers, crm_company_id, module = "payroll-recorder", mappings, action } = body;
    
    const supabase = getSupabaseClient();
    
    // ========================================================================
    // Action: DEBUG - Check database connectivity and saved mappings
    // ========================================================================
    if (action === "debug") {
      const debugInfo: Record<string, unknown> = {
        version: VERSION,
        crm_company_id,
        module,
        headers_count: headers?.length || 0,
        supabase_url: SUPABASE_URL,
        has_service_key: !!SUPABASE_SERVICE_KEY
      };
      
      // Try to fetch saved mappings
      try {
        const { data, error, count } = await supabase
          .from("ada_customer_column_mappings")
          .select("*", { count: "exact" })
          .eq("crm_company_id", crm_company_id || "")
          .eq("module", module);
        
        debugInfo.saved_mappings_error = error?.message || null;
        debugInfo.saved_mappings_count = count;
        debugInfo.saved_mappings_sample = data?.slice(0, 5) || [];
      } catch (e) {
        debugInfo.saved_mappings_exception = String(e);
      }
      
      // Try to fetch dictionary
      try {
        const { data, error, count } = await supabase
          .from("ada_payroll_column_dictionary")
          .select("*", { count: "exact" })
          .eq("module", module)
          .limit(3);
        
        debugInfo.dictionary_error = error?.message || null;
        debugInfo.dictionary_count = count;
        debugInfo.dictionary_sample = data?.slice(0, 3) || [];
      } catch (e) {
        debugInfo.dictionary_exception = String(e);
      }
      
      return new Response(
        JSON.stringify(debugInfo, null, 2),
        { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: ANALYZE
    // ========================================================================
    if (action === "analyze") {
      if (!headers || !Array.isArray(headers) || headers.length === 0) {
        return new Response(
          JSON.stringify({ error: "Headers array is required" }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      console.log(`[Analyze] ${headers.length} headers, company: ${crm_company_id || "none"}, module: ${module}`);
      
      // Filter out blocked headers entirely - they won't appear in results
      // NOTE: Header filtering is now database-driven via include_in_matrix
      // All headers are passed through; client filters based on include_in_matrix
      const filteredHeaders = headers;
      
      const resultMappings: MappingResult[] = [];
      const processedHeaders = new Set<string>();
      
      // =====================================================================
      // Priority 1: Saved company mappings (these always win)
      // =====================================================================
      let savedMappings = new Map<string, MappingResult>();
      if (crm_company_id) {
        savedMappings = await getSavedMappings(supabase, crm_company_id, module, filteredHeaders);
        for (const [header, mapping] of savedMappings) {
          // NOTE: BLOCKED_TARGETS check removed - include_in_matrix field handles this
          resultMappings.push(mapping);
          processedHeaders.add(header);
        }
      }
      
      // =====================================================================
      // Get ALL matches from both dictionaries for remaining headers
      // =====================================================================
      const remainingHeaders = filteredHeaders.filter(h => !processedHeaders.has(h));
      
      const amountMatches = await getAmountMatches(supabase, module, remainingHeaders);
      const dimensionMatches = await getDimensionMatches(supabase, remainingHeaders);
      
      // =====================================================================
      // Process each remaining header with ambiguity detection
      // =====================================================================
      for (const header of remainingHeaders) {
        const hasAmount = amountMatches.has(header);
        const hasDimension = dimensionMatches.has(header);
        
        // CASE: Both amount AND dimension match → AMBIGUOUS
        if (hasAmount && hasDimension) {
          const amountMatch = amountMatches.get(header)!;
          const dimensionMatch = dimensionMatches.get(header)!;
          
          // NOTE: BLOCKED_TARGETS check removed - include_in_matrix handles filtering
          // Show ambiguity and let user/database decide
          resultMappings.push({
            raw_header: header,
            kind: "ambiguous",
            target: null,
            source: "ambiguous",
            confidence: 0.5,
            amount_option: {
              target: amountMatch.target,      // pf_column_name
              confidence: amountMatch.confidence
            },
            dimension_option: {
              target: dimensionMatch.target,   // normalized_dimension
              confidence: dimensionMatch.confidence
            }
          });
          processedHeaders.add(header);
          continue;
        }
        
        // CASE: Only amount match
        if (hasAmount) {
          const match = amountMatches.get(header)!;
          
          // NOTE: BLOCKED_TARGETS check removed - include_in_matrix handles filtering
          resultMappings.push({
            raw_header: header,
            kind: "amount",
            target: match.target,       // pf_column_name
            source: match.isExact ? "amount" : "fuzzy",
            confidence: match.confidence
          });
          processedHeaders.add(header);
          continue;
        }
        
        // CASE: Only dimension match
        if (hasDimension) {
          const match = dimensionMatches.get(header)!;
          
          // NOTE: BLOCKED_TARGETS check removed - include_in_matrix handles filtering
          resultMappings.push({
            raw_header: header,
            kind: "dimension",
            target: match.target,       // normalized_dimension
            source: match.isExact ? "dimension" : "fuzzy",
            confidence: match.confidence
          });
          processedHeaders.add(header);
          continue;
        }
        
        // CASE: No match → unmapped
        resultMappings.push({
          raw_header: header,
          kind: null,
          target: null,
          source: "unmapped",
          confidence: 0
        });
      }
      
      // =====================================================================
      // Enrich amount mappings with GL account data
      // =====================================================================
      if (crm_company_id) {
        const targets = resultMappings
          .filter(m => m.target && m.kind === "amount")
          .map(m => m.target as string);
        
        if (targets.length > 0) {
          const glMappings = await getGLMappings(supabase, crm_company_id, module, targets);
          
          for (const mapping of resultMappings) {
            if (mapping.target && mapping.kind === "amount" && glMappings.has(mapping.target)) {
              const gl = glMappings.get(mapping.target)!;
              mapping.gl_account = gl.gl_account;
              mapping.gl_account_name = gl.gl_account_name;
            }
          }
        }
      }
      
      // =====================================================================
      // Apply dimension-specific overrides (e.g., Department/Department Description)
      // =====================================================================
      applyDimensionOverrides(resultMappings);
      
      // =====================================================================
      // Build summary stats
      // =====================================================================
      const amounts = resultMappings.filter(m => m.kind === "amount").length;
      const dimensions = resultMappings.filter(m => m.kind === "dimension").length;
      const ambiguous = resultMappings.filter(m => m.kind === "ambiguous").length;
      const unmapped = resultMappings.filter(m => m.source === "unmapped").length;
      const fuzzyCount = resultMappings.filter(m => m.source === "fuzzy").length;
      const withGL = resultMappings.filter(m => m.gl_account).length;
      const matched = amounts + dimensions;
      
      console.log(`[Analyze] Result: ${matched} matched (${amounts} amount, ${dimensions} dimension), ${fuzzyCount} fuzzy, ${ambiguous} ambiguous, ${unmapped} unmapped, ${withGL} with GL`);
      
      return new Response(
        JSON.stringify({
          mappings: resultMappings,
          source: savedMappings.size > 0 ? "saved" : "dictionary",
          matched,
          unmapped,
          amounts,
          dimensions,
          ambiguous,
          fuzzy: fuzzyCount,
          with_gl: withGL,
          total: headers.length
        }),
        { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: SAVE
    // ========================================================================
    if (action === "save") {
      if (!crm_company_id) {
        return new Response(
          JSON.stringify({ error: "crm_company_id is required for saving mappings" }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      if (!mappings || !Array.isArray(mappings)) {
        return new Response(
          JSON.stringify({ error: "mappings array is required" }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      console.log(`[Save] Saving ${mappings.length} mappings for company: ${crm_company_id}, module: ${module}`);
      
      const result = await saveMappings(supabase, crm_company_id, module, mappings);
      
      return new Response(
        JSON.stringify(result),
        { status: result.success ? 200 : 207, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: GET_OPTIONS - Return dictionary options for dropdowns
    // ========================================================================
    if (action === "get_options") {
      console.log(`[GetOptions] Fetching options for module: ${module}`);
      
      // Get amount options from ada_payroll_column_dictionary
      // Note: Filtering for expense_review_include handled client-side via include_in_matrix
      const { data: amountData, error: amountError } = await supabase
        .from("ada_payroll_column_dictionary")
        .select("pf_column_name")
        .eq("module", module)
        .not("pf_column_name", "is", null);
      
      if (amountError) {
        console.error("[GetOptions] Error fetching amounts:", amountError);
      }
      
      // Get dimension options from ada_payroll_dimensions
      const { data: dimensionData, error: dimensionError } = await supabase
        .from("ada_payroll_dimensions")
        .select("normalized_dimension")
        .not("normalized_dimension", "is", null);
      
      if (dimensionError) {
        console.error("[GetOptions] Error fetching dimensions:", dimensionError);
      }
      
      // Extract unique values and sort
      const amountOptions = [...new Set(
        (amountData || []).map(r => r.pf_column_name).filter(Boolean)
      )].sort();
      
      const dimensionOptions = [...new Set(
        (dimensionData || []).map(r => r.normalized_dimension).filter(Boolean)
      )].sort();
      
      console.log(`[GetOptions] Found ${amountOptions.length} amounts, ${dimensionOptions.length} dimensions`);
      
      return new Response(
        JSON.stringify({
          amount_options: amountOptions,
          dimension_options: dimensionOptions,
          module
        }),
        { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: GET_EXPENSE_TAXONOMY - Return full taxonomy for expense review
    // ========================================================================
    if (action === "get_expense_taxonomy") {
      console.log(`[GetExpenseTaxonomy] Fetching taxonomy for module: ${module}`);
      
      // Get ALL measures from ada_payroll_column_dictionary with expense review fields
      // Load all rows for module='payroll-recorder' without filtering by term_type/side/mapping_type
      // This ensures benefits, fees, and all other column types are included
      // pf_column_name is the canonical key used by PR_Data_Clean headers
      // 
      // INCLUSION LOGIC (using `side` as primary signal):
      // - side = 'er' (employer) → include in expense review totals
      // - side = 'ee' (employee) → exclude from totals
      // - side = 'na' (not applicable) → exclude from totals
      // expense_review_bucket is only for grouping (FIXED/VARIABLE/BURDEN), not inclusion
      const { data: measureData, error: measureError } = await supabase
        .from("ada_payroll_column_dictionary")
        .select(`
          pf_column_name,
          side,
          expense_review_bucket,
          expense_review_include,
          default_sign,
          display_order
        `)
        .eq("module", module)
        .not("pf_column_name", "is", null);
      
      if (measureError) {
        console.error("[GetExpenseTaxonomy] Error fetching measures:", measureError);
      }
      
      // Debug logging: show raw row count and check for specific columns
      console.log(`[GetExpenseTaxonomy] Raw measureData row count: ${measureData?.length || 0}`);
      const checkCols = ["401k_employer_amount", "fees_peo_employer_amount", "401K_Employer_Amount", "Fees_PEO_Employer_Amount"];
      checkCols.forEach(col => {
        const found = measureData?.find(r => r.pf_column_name?.toLowerCase() === col.toLowerCase());
        console.log(`[GetExpenseTaxonomy] Contains ${col}: ${found ? 'YES' : 'NO'}${found ? ` (side=${found.side}, bucket=${found.expense_review_bucket})` : ''}`);
      });
      
      // Get dimensions from ada_payroll_dimensions
      const { data: dimensionData, error: dimensionError } = await supabase
        .from("ada_payroll_dimensions")
        .select(`
          normalized_dimension,
          semantic_group
        `)
        .not("normalized_dimension", "is", null);
      
      if (dimensionError) {
        console.error("[GetExpenseTaxonomy] Error fetching dimensions:", dimensionError);
      }
      
      console.log(`[GetExpenseTaxonomy] Raw dimensionData row count: ${dimensionData?.length || 0}`);
      
      // Build measure lookup map (canonical header -> metadata)
      // A column is "classified" if it exists in the dictionary
      // 
      // INCLUSION LOGIC (using `side` as PRIMARY signal):
      // - side = 'er' (employer) → include = true
      // - side = 'ee' (employee) → include = false
      // - side = 'na' (not applicable) → include = false
      // - side = null → fallback to expense_review_include field
      // 
      // expense_review_bucket is only for grouping/presentation (FIXED/VARIABLE/BURDEN)
      // All employer-paid items (side='er') MUST be included, especially for PEO scenarios
      const measures: Record<string, {
        bucket: string | null;
        include: boolean;
        sign: number;
        displayOrder: number;
        side: string | null;
      }> = {};
      
      (measureData || []).forEach(row => {
        if (row.pf_column_name) {
          // Normalize side to lowercase, handle null/blank safely
          const side = (row.side || "").toString().toLowerCase().trim() || null;
          
          // INCLUSION LOGIC (permissive by default):
          // - side='ee' → EXCLUDE (employee deductions)
          // - side='na' → EXCLUDE (summary/info)
          // - expense_review_include=false → EXCLUDE
          // - Everything else → INCLUDE (er, null, undefined, blank)
          // This ensures columns without metadata default to INCLUDED
          let include: boolean;
          
          if (side === 'ee') {
            include = false;  // Employee-paid → exclude
          } else if (side === 'na') {
            include = false;  // Summary/info → exclude
          } else if (row.expense_review_include === false) {
            include = false;  // Explicit exclude flag
          } else {
            include = true;   // Everything else included (er, null, blank)
          }
          
          measures[row.pf_column_name] = {
            bucket: row.expense_review_bucket || null,
            include,
            sign: row.default_sign ?? 1,
            displayOrder: row.display_order ?? 100,
            side
          };
        }
      });
      
      // Build dimension set (canonical headers that are dimensions)
      const dimensions = new Set(
        (dimensionData || []).map(r => r.normalized_dimension).filter(Boolean)
      );
      
      console.log(`[GetExpenseTaxonomy] Built ${Object.keys(measures).length} measures, ${dimensions.size} dimensions`);
      
      return new Response(
        JSON.stringify({
          measures,
          dimensions: Array.from(dimensions),
          module,
          debug: {
            measureRowCount: measureData?.length || 0,
            dimensionRowCount: dimensionData?.length || 0
          }
        }),
        { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: GET_DIMENSIONS - Return dimension mappings for a provider
    // ========================================================================
    if (action === "get_dimensions") {
      const { provider } = body;
      
      console.log(`[GetDimensions] Fetching dimensions for provider: ${provider || "all"}`);
      
      try {
        let query = supabase
          .from("ada_payroll_dimensions")
          .select("raw_term, normalized_dimension");
        
        // Filter by provider if specified
        if (provider) {
          query = query.eq("provider", provider);
        }
        
        const { data, error } = await query;
        
        if (error) {
          console.error("[GetDimensions] Error fetching dimensions:", error);
          return new Response(
            JSON.stringify({ success: false, error: error.message }),
            { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
        
        console.log(`[GetDimensions] Found ${data?.length || 0} dimension mappings`);
        
        return new Response(
          JSON.stringify({
            success: true,
            dimensions: data || [],
            count: data?.length || 0
          }),
          { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      } catch (e) {
        console.error("[GetDimensions] Exception:", e);
        return new Response(
          JSON.stringify({ success: false, error: "Failed to fetch dimensions" }),
          { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
    }
    
    // ========================================================================
    // Action: VALIDATE_EMAIL - Check email authorization for installation
    // ========================================================================
    if (action === "validate_email") {
      const { email, installation_key } = body;
      
      if (!email || !installation_key) {
        return new Response(
          JSON.stringify({ 
            success: false, 
            authorized: false,
            error: "Email and installation key required" 
          }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      console.log(`[ValidateEmail] Checking authorization for: ${email}`);
      
      // Step 1: Get crm_company_id from installation
      const { data: installation, error: installError } = await supabase
        .from("ada_addin_installations")
        .select("crm_company_id")
        .eq("installation_key", installation_key)
        .single();
      
      if (installError || !installation) {
        console.log("[ValidateEmail] Invalid installation key");
        return new Response(
          JSON.stringify({ 
            success: false, 
            authorized: false,
            error: "Invalid installation" 
          }),
          { status: 404, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      const { crm_company_id } = installation;
      console.log(`[ValidateEmail] Installation company: ${crm_company_id}`);
      
      // Step 2: Check if email exists in contacts for this company
      const { data: contact, error: contactError } = await supabase
        .from("admin_crm_contacts")
        .select("id, related_user_id")
        .eq("email", email.toLowerCase())
        .eq("crm_company_id", crm_company_id)
        .maybeSingle();
      
      // If found in contacts for this company, authorized!
      if (contact) {
        console.log("[ValidateEmail] ✓ Authorized as company contact");
        return new Response(
          JSON.stringify({ 
            success: true, 
            authorized: true,
            reason: "company_contact"
          }),
          { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      // Step 3: Check if user is an admin (bypass company check)
      const { data: contactAnyCompany, error: anyContactError } = await supabase
        .from("admin_crm_contacts")
        .select("related_user_id")
        .eq("email", email.toLowerCase())
        .maybeSingle();
      
      if (contactAnyCompany && contactAnyCompany.related_user_id) {
        // Check if this user has admin role
        const { data: userRole, error: roleError } = await supabase
          .from("user_roles")
          .select("role")
          .eq("user_id", contactAnyCompany.related_user_id)
          .eq("role", "admin")
          .maybeSingle();
        
        if (userRole) {
          console.log("[ValidateEmail] ✓ Authorized as admin user");
          return new Response(
            JSON.stringify({ 
              success: true, 
              authorized: true,
              reason: "admin_user"
            }),
            { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
      }
      
      // Step 4: Not authorized
      console.log("[ValidateEmail] ✗ Not authorized");
      return new Response(
        JSON.stringify({ 
          success: true, 
          authorized: false,
          reason: "unauthorized"
        }),
        { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    
    // ========================================================================
    // Action: BOOTSTRAP - Auto-load company metadata from installation key
    // ========================================================================
    if (action === "bootstrap") {
      const { installation_key } = body;
      
      if (!installation_key) {
        return new Response(
          JSON.stringify({ success: false, error: "installation_key is required" }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      console.log(`[Bootstrap] Looking up installation: ${installation_key.substring(0, 12)}...`);
      
      try {
        const { data, error } = await supabase
          .from("ada_addin_installations")
          .select("crm_company_id, ss_company_name, ss_accounting_software, pto_payroll_provider, pr_payroll_provider")
          .eq("installation_key", installation_key)
          .single();
        
        if (error || !data) {
          console.log("[Bootstrap] Installation key not found");
          return new Response(
            JSON.stringify({ success: false, error: "Installation not found" }),
            { status: 404, headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
        
        console.log(`[Bootstrap] Found company: ${data.ss_company_name} (${data.crm_company_id})`);
        
        return new Response(
          JSON.stringify({
            success: true,
            crm_company_id: data.crm_company_id,
            ss_company_name: data.ss_company_name,
            ss_accounting_software: data.ss_accounting_software,
            pto_payroll_provider: data.pto_payroll_provider,
            pr_payroll_provider: data.pr_payroll_provider
          }),
          { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      } catch (e) {
        console.error("[Bootstrap] Error:", e);
        return new Response(
          JSON.stringify({ success: false, error: "Bootstrap lookup failed" }),
          { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
    }
    
    return new Response(
      JSON.stringify({ error: "Invalid action. Use 'analyze', 'save', 'get_options', 'get_expense_taxonomy', 'get_dimensions', or 'bootstrap'" }),
      { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
    
  } catch (error) {
    console.error("[ColumnMapper] Error:", error);
    return new Response(
      JSON.stringify({ error: "Internal server error" }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
