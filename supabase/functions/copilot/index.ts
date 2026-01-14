/**
 * Ada - Prairie Forge AI Assistant
 * Supabase Edge Function powered by Claude (Anthropic)
 * 
 * Features:
 * - Fetches system prompts from database (admin-editable)
 * - Logs all conversations for debugging and improvement
 * - Supports multiple prompt personalities
 * 
 * COST ESTIMATES (Claude 3.5 Sonnet):
 * - Input: $3/million tokens (~$0.003 per 1K tokens)
 * - Output: $15/million tokens (~$0.015 per 1K tokens)
 * - Typical question: ~$0.01-0.03
 * - 100 questions/day ≈ $1-3/day (cheaper than GPT-4!)
 */

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

declare const Deno: {
  env: {
    get: (key: string) => string | undefined;
  };
};

// Configuration
const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") || "https://jgciqwzwacaesqjaoadc.supabase.co";
const SUPABASE_SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

// Default model configuration (can be overridden by database)
// Using latest Claude Sonnet 4 as of Jan 2026
const DEFAULT_MODEL = "claude-sonnet-4-20250514"; 
const DEFAULT_MAX_TOKENS = 2048; // Claude supports up to 8192
const DEFAULT_TEMPERATURE = 0.7;

// Logging / privacy controls
const STORE_AI_RESPONSES = Deno.env.get("STORE_AI_RESPONSES") === "true";
// "keys" stores only Object.keys(context); "full" stores the full context payload
const STORE_CONTEXT_MODE = (Deno.env.get("STORE_CONTEXT_MODE") || "keys").toLowerCase();

// Multi-tenant scoping controls
// If enabled, queries will filter by customer_id when provided.
// IMPORTANT: Only enable this if your tables have a customer_id column.
const ENABLE_CUSTOMER_SCOPING = Deno.env.get("ENABLE_CUSTOMER_SCOPING") === "true";

// CORS controls
// Comma-separated allowlist of origins. If empty/unset, falls back to "*".
const ALLOWED_ORIGINS = (Deno.env.get("ALLOWED_ORIGINS") || "").split(",").map((s: string) => s.trim()).filter(Boolean);

const allowedModules = new Set(["payroll-recorder", "pto-accrual", "module-selector", "global"]);
const allowedFunctions = new Set(["mapping", "analysis", "validation"]);

function buildCorsHeaders(origin: string | null) {
  const allowAll = ALLOWED_ORIGINS.length === 0;
  const isAllowed = allowAll || (!!origin && ALLOWED_ORIGINS.includes(origin));
  const allowOrigin = allowAll ? "*" : (isAllowed ? origin! : "null");
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Vary": "Origin",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };
}

function isOriginAllowed(origin: string | null) {
  if (ALLOWED_ORIGINS.length === 0) return true;
  return !!origin && ALLOWED_ORIGINS.includes(origin);
}

function sanitizeModuleKey(input: string | null | undefined) {
  if (!input) return null;
  const key = String(input).trim();
  if (!key) return null;
  return allowedModules.has(key) ? key : null;
}

function sanitizeFunctionContext(input: string | null | undefined) {
  const key = String(input || "analysis").trim();
  return allowedFunctions.has(key) ? key : "analysis";
}

function sanitizeStoredContext(context: Record<string, unknown> | undefined) {
  if (!context) return null;
  if (STORE_CONTEXT_MODE === "full") return context;
  // Default: store only keys to reduce risk of sensitive content being logged
  return Object.keys(context);
}

interface AdaRequest {
  prompt: string;
  context?: Record<string, unknown>;
  systemPrompt?: string;
  promptName?: string; // Name of the system prompt to use from database
  module?: string; // Module key: 'payroll-recorder', 'pto-accrual'
  function?: string; // Function context: 'mapping', 'analysis', 'validation'
  history?: Array<{ role: string; content: string }>;
  sessionId?: string;
  customerId?: string;
}

async function fetchNamedSystemPrompt(promptName: string, customerId: string | null): Promise<{ system_prompt: string; model: string; max_tokens: number; temperature: number } | null> {
  const supabase = getSupabaseClient();
  if (!supabase) return null;

  try {
    let query = supabase
      .from('ada_system_prompts')
      .select('prompt_text, model, max_tokens, temperature')
      .eq('name', promptName)
      .eq('is_active', true);

    if (ENABLE_CUSTOMER_SCOPING) {
      if (!customerId) {
        console.warn("[Ada] ENABLE_CUSTOMER_SCOPING is true but request.customerId is missing");
      } else {
        query = query.or(`customer_id.is.null,customer_id.eq.${customerId}`);
      }
    }

    const { data, error } = await query.single();
    if (error || !data) return null;
    return {
      system_prompt: data.prompt_text,
      model: data.model || DEFAULT_MODEL,
      max_tokens: data.max_tokens || DEFAULT_MAX_TOKENS,
      temperature: data.temperature || DEFAULT_TEMPERATURE,
    };
  } catch (e) {
    console.error('Failed to fetch named system prompt:', e);
    return null;
  }
}

interface ModuleConfigResult {
  system_prompt: string | null;
  welcome_message: string | null;
  model: string;
  max_tokens: number;
  temperature: number;
  ada_context_mapping: string | null;
  ada_context_analysis: string | null;
  ada_context_validation: string | null;
}

// Create Supabase client for database operations
function getSupabaseClient() {
  if (!SUPABASE_SERVICE_KEY) {
    console.warn("SUPABASE_SERVICE_ROLE_KEY not set, database features disabled");
    return null;
  }
  return createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
}

// Fetch module config from database (unified table with prompts)
async function fetchModuleConfig(moduleKey: string, customerId: string | null): Promise<ModuleConfigResult | null> {
  const supabase = getSupabaseClient();
  if (!supabase) return null;

  try {
    let query = supabase
      .from('ada_module_config')
      .select('system_prompt, welcome_message, model, max_tokens, temperature, ada_context_mapping, ada_context_analysis, ada_context_validation')
      .eq('module_key', moduleKey)
      .eq('is_active', true);

    if (ENABLE_CUSTOMER_SCOPING) {
      if (!customerId) {
        console.warn("[Ada] ENABLE_CUSTOMER_SCOPING is true but request.customerId is missing");
      } else {
        query = query.or(`customer_id.is.null,customer_id.eq.${customerId}`);
      }
    }

    const { data, error } = await query.single();

    if (error) {
      console.log('No module config found for:', moduleKey);
      return null;
    }

    return data;
  } catch (e) {
    console.error('Failed to fetch module config:', e);
    return null;
  }
}

// Fetch global/default config as fallback
async function fetchGlobalConfig(customerId: string | null): Promise<ModuleConfigResult | null> {
  const supabase = getSupabaseClient();
  if (!supabase) return null;

  try {
    // Try 'global' first, then 'default'
    let query = supabase
      .from('ada_module_config')
      .select('system_prompt, welcome_message, model, max_tokens, temperature, ada_context_mapping, ada_context_analysis, ada_context_validation')
      .eq('module_key', 'global')
      .eq('is_active', true);

    if (ENABLE_CUSTOMER_SCOPING) {
      if (!customerId) {
        console.warn("[Ada] ENABLE_CUSTOMER_SCOPING is true but request.customerId is missing");
      } else {
        query = query.or(`customer_id.is.null,customer_id.eq.${customerId}`);
      }
    }

    let { data, error } = await query.single();

    if (error || !data) {
      // Fallback: try ada_system_prompts for backward compatibility
      const { data: legacyData, error: legacyError } = await supabase
        .from('ada_system_prompts')
        .select('prompt_text, model, max_tokens, temperature')
        .eq('name', 'default')
        .eq('is_active', true)
        .single();
      
      if (!legacyError && legacyData) {
        return {
          system_prompt: legacyData.prompt_text,
          welcome_message: null,
          model: legacyData.model || DEFAULT_MODEL,
          max_tokens: legacyData.max_tokens || DEFAULT_MAX_TOKENS,
          temperature: legacyData.temperature || DEFAULT_TEMPERATURE,
          ada_context_mapping: null,
          ada_context_analysis: null,
          ada_context_validation: null,
        };
      }
      return null;
    }

    return data;
  } catch (e) {
    console.error('Failed to fetch global config:', e);
    return null;
  }
}

// Fetch relevant knowledge sources from database
async function fetchKnowledgeSources(moduleKey: string | null, functionContext: string | null, customerId: string | null): Promise<string> {
  const supabase = getSupabaseClient();
  if (!supabase) return '';

  try {
    let query = supabase
      .from('ada_knowledge_sources')
      .select('source_type, title, content')
      .eq('is_active', true)
      .order('priority', { ascending: false })
      .limit(5);

    if (ENABLE_CUSTOMER_SCOPING) {
      if (!customerId) {
        console.warn("[Ada] ENABLE_CUSTOMER_SCOPING is true but request.customerId is missing");
      } else {
        query = query.or(`customer_id.is.null,customer_id.eq.${customerId}`);
      }
    }

    // Filter by module (include global entries where module_key is null)
    if (moduleKey) {
      query = query.or(`module_key.is.null,module_key.eq.${moduleKey}`);
    }

    // Filter by function context (include global entries where function_context is null)
    if (functionContext) {
      query = query.or(`function_context.is.null,function_context.eq.${functionContext}`);
    }

    const { data, error } = await query;

    if (error || !data?.length) {
      return '';
    }

    // Format knowledge sources for injection
    const knowledgeBlock = data.map(k => 
      `### [${k.source_type.toUpperCase()}] ${k.title}\n${k.content}`
    ).join('\n\n');

    return `\n\n## REFERENCE KNOWLEDGE\nUse this knowledge to answer user questions when relevant:\n\n${knowledgeBlock}`;
  } catch (e) {
    console.error('Failed to fetch knowledge sources:', e);
    return '';
  }
}

// Log conversation to database
async function logConversation(
  request: AdaRequest,
  response: string | null,
  tokensUsed: number | null,
  latencyMs: number,
  error: string | null,
  model: string,
  moduleContext: string | null,
  functionContext: string
) {
  const supabase = getSupabaseClient();
  if (!supabase) return;

  try {
    await supabase
      .from('ada_conversations')
      .insert({
        session_id: request.sessionId || null,
        crm_company_id: request.customerId || null, // Renamed from customer_id to crm_company_id
        prompt_name: moduleContext, // Table uses 'prompt_name' not 'module_context'
        user_prompt: request.prompt,
        context: sanitizeStoredContext(request.context),
        ai_response: STORE_AI_RESPONSES ? response : null,
        model: model,
        tokens_used: tokensUsed,
        latency_ms: latencyMs,
        error: error,
      });
  } catch (e) {
    console.error('Failed to log conversation:', e);
  }
}

serve(async (req) => {
  const startTime = Date.now();

  const origin = req.headers.get("Origin");
  const corsHeaders = buildCorsHeaders(origin);
  
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    if (!isOriginAllowed(origin)) {
      return new Response("forbidden", { status: 403, headers: corsHeaders });
    }
    return new Response("ok", { headers: corsHeaders });
  }

  if (!isOriginAllowed(origin)) {
    return new Response(
      JSON.stringify({ error: "Origin not allowed", code: "CORS_FORBIDDEN" }),
      { status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }

  let body: AdaRequest | null = null;
  let model = DEFAULT_MODEL;

  try {
    // Validate API key is configured
    if (!ANTHROPIC_API_KEY) {
      console.error("ANTHROPIC_API_KEY not configured");
      return new Response(
        JSON.stringify({ 
          error: "Ada is not configured yet. Please contact Prairie Forge support.",
          code: "CONFIG_ERROR"
        }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Parse request
    body = await req.json();
    if (!body) {
      return new Response(
        JSON.stringify({ error: "Invalid request body", code: "INVALID_REQUEST" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    const { prompt, context, systemPrompt, promptName, history } = body;
    const moduleKey = sanitizeModuleKey(body.module);
    const functionContext = sanitizeFunctionContext(body.function);
    const customerId = body.customerId || null;

    if (!prompt?.trim()) {
      return new Response(
        JSON.stringify({ error: "Please ask Ada a question!", code: "INVALID_REQUEST" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Build the effective system prompt from unified ada_module_config
    let effectiveSystemPrompt = systemPrompt;
    let maxTokens = DEFAULT_MAX_TOKENS;
    let temperature = DEFAULT_TEMPERATURE;

    if (!systemPrompt) {
      // Optional: select a named prompt (legacy table)
      if (promptName) {
        const named = await fetchNamedSystemPrompt(promptName, customerId);
        if (named) {
          effectiveSystemPrompt = named.system_prompt;
          // Only use database model if it's a valid Claude model
          if (named.model && named.model.startsWith('claude-')) {
            model = named.model;
          } else {
            model = DEFAULT_MODEL;
          }
          maxTokens = named.max_tokens;
          temperature = named.temperature;
        }
      }

      // Get config from module or fall back to global
      let config: ModuleConfigResult | null = null;
      
      if (!effectiveSystemPrompt && moduleKey) {
        config = await fetchModuleConfig(moduleKey, customerId);
      }
      
      // Fall back to global config if no module-specific config
      if (!effectiveSystemPrompt && !config) {
        config = await fetchGlobalConfig(customerId);
      }
      
      if (config) {
        effectiveSystemPrompt = config.system_prompt || '';
        // Only use database model if it's set AND starts with "claude-"
        // This prevents using outdated OpenAI models or deprecated Claude versions
        if (config.model && config.model.startsWith('claude-')) {
          model = config.model;
        } else {
          model = DEFAULT_MODEL; // Use code default for safety
        }
        maxTokens = config.max_tokens || DEFAULT_MAX_TOKENS;
        temperature = config.temperature || DEFAULT_TEMPERATURE;
        
        // Append function-specific context
        const contextAdditions: Record<string, string | null> = {
          'mapping': config.ada_context_mapping,
          'analysis': config.ada_context_analysis,
          'validation': config.ada_context_validation,
        };
        const additionalContext = contextAdditions[functionContext];
        if (additionalContext && effectiveSystemPrompt) {
          effectiveSystemPrompt += `\n\n## ADDITIONAL CONTEXT FOR ${functionContext.toUpperCase()}\n${additionalContext}`;
        }
      }
      
      // Inject relevant knowledge sources
      const knowledgeBlock = await fetchKnowledgeSources(moduleKey, functionContext, customerId);
      if (knowledgeBlock && effectiveSystemPrompt) {
        effectiveSystemPrompt += knowledgeBlock;
      }
    }

    // Build messages for Claude (different format than OpenAI)
    const { system, messages } = buildClaudeMessages(prompt, context, effectiveSystemPrompt, history);

    // Call Anthropic Claude API
    const anthropicResponse = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model,
        max_tokens: maxTokens,
        temperature,
        system, // Claude uses separate system parameter
        messages,
      }),
    });

    if (!anthropicResponse.ok) {
      const errorData = await anthropicResponse.json();
      console.error("Anthropic API error:", errorData);
      
      const latencyMs = Date.now() - startTime;
      await logConversation(body, null, null, latencyMs, `Anthropic error: ${anthropicResponse.status}`, model, moduleKey, functionContext);
      
      if (anthropicResponse.status === 429) {
        return new Response(
          JSON.stringify({ 
            error: "Ada is thinking hard right now. Please try again in a moment!",
            code: "AI_BUSY"
          }),
          { status: 429, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      
      return new Response(
        JSON.stringify({ 
          error: "Ada encountered an issue. Please try again.",
          code: "AI_ERROR"
        }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const completion = await anthropicResponse.json();
    
    // Claude response format is different from OpenAI
    const responseMessage = completion.content?.[0]?.text || "I couldn't generate a response. Please try rephrasing your question.";
    
    // Calculate total tokens (input + output)
    const tokensUsed = (completion.usage?.input_tokens || 0) + (completion.usage?.output_tokens || 0);
    const latencyMs = Date.now() - startTime;

    // Log successful conversation
    await logConversation(body, responseMessage, tokensUsed, latencyMs, null, model, moduleKey, functionContext);

    console.log(`Ada responded: ${tokensUsed} tokens used (${completion.usage?.input_tokens} in, ${completion.usage?.output_tokens} out), ${latencyMs}ms, module: ${moduleKey || 'none'}, function: ${functionContext}`);

    // Return successful response
    return new Response(
      JSON.stringify({
        message: responseMessage,
        usage: {
          tokens: tokensUsed,
          inputTokens: completion.usage?.input_tokens || 0,
          outputTokens: completion.usage?.output_tokens || 0,
          model: model,
          latencyMs: latencyMs,
          module: moduleKey,
          function: functionContext
        }
      }),
      { 
        status: 200, 
        headers: { ...corsHeaders, "Content-Type": "application/json" } 
      }
    );

  } catch (error) {
    console.error("Ada function error:", error);
    
    const latencyMs = Date.now() - startTime;
    if (body) {
      const moduleKey = sanitizeModuleKey(body.module);
      const functionContext = sanitizeFunctionContext(body.function);
      await logConversation(body, null, null, latencyMs, String(error), model, moduleKey, functionContext);
    }
    
    return new Response(
      JSON.stringify({ 
        error: "Something went wrong. Please try again!",
        code: "INTERNAL_ERROR"
      }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});

/**
 * Build the messages array for Claude
 * Claude has a different format than OpenAI:
 * - System prompt is a separate parameter
 * - Messages alternate between user and assistant
 * - No "system" role in messages array
 */
function buildClaudeMessages(
  prompt: string, 
  context: Record<string, unknown> | undefined,
  systemPrompt: string | undefined,
  history: Array<{ role: string; content: string }> | undefined
): { system: string; messages: Array<{ role: "user" | "assistant"; content: string }> } {
  
  // Ada's default personality (fallback if no database prompt)
  const defaultSystemPrompt = `You are Ada, Prairie Forge's expert financial analyst assistant. You're embedded in an Excel add-in helping accountants and CFOs review payroll data.

Your personality:
- Warm, professional, and confident
- You explain complex financial concepts simply
- You're proactive about spotting issues
- You celebrate wins and acknowledge good data

Your expertise:
- Payroll expense analysis and validation
- Trend identification and variance analysis
- Executive-ready insights and talking points
- Journal entry preparation and validation

Communication style:
- Start with a brief, direct answer
- Use bullet points for clarity
- Highlight issues with ⚠️ and successes with ✓
- Format currency as $X,XXX (no unnecessary decimals)
- Format percentages as X.X%
- End with 2-3 actionable next steps

When given spreadsheet context, reference specific numbers from the data.
Be confident in your analysis but acknowledge when data is limited.`;

  const system = systemPrompt || defaultSystemPrompt;
  
  const messages: Array<{ role: "user" | "assistant"; content: string }> = [];

  // Add context as first user message if provided
  if (context && Object.keys(context).length > 0) {
    const contextSummary = formatContextForAI(context);
    messages.push({
      role: "user",
      content: `Here's the current spreadsheet data:\n${contextSummary}`
    });
    
    // Claude requires alternating messages, so add a brief acknowledgment
    messages.push({
      role: "assistant",
      content: "I've reviewed the spreadsheet data. What would you like to know?"
    });
  }

  // Add conversation history (last 8 messages for good context)
  if (history?.length) {
    const recentHistory = history.slice(-8);
    for (const msg of recentHistory) {
      if (msg.role === "user" || msg.role === "assistant") {
        messages.push({ 
          role: msg.role as "user" | "assistant", 
          content: msg.content 
        });
      }
    }
  }

  // Add current prompt
  messages.push({ role: "user", content: prompt });

  return { system, messages };
}

/**
 * Format context data for AI consumption
 */
function formatContextForAI(context: Record<string, unknown>): string {
  const parts: string[] = [];

  if (context.period) {
    parts.push(`Period: ${context.period}`);
  }

  if (context.summary) {
    const s = context.summary as Record<string, unknown>;
    parts.push(`Summary:`);
    if (s.total) parts.push(`  - Total Payroll: $${Number(s.total).toLocaleString()}`);
    if (s.employeeCount) parts.push(`  - Employee Count: ${s.employeeCount}`);
    if (s.avgPerEmployee) parts.push(`  - Avg/Employee: $${Number(s.avgPerEmployee).toLocaleString()}`);
  }

  if (context.departments && Array.isArray(context.departments)) {
    parts.push(`\nDepartment Breakdown:`);
    for (const dept of context.departments.slice(0, 8)) {
      const d = dept as Record<string, unknown>;
      const pct = d.percentOfTotal ? ` (${(Number(d.percentOfTotal) * 100).toFixed(1)}%)` : '';
      parts.push(`  - ${d.name}: $${Number(d.total).toLocaleString()}${pct}`);
    }
  }

  if (context.journalEntry) {
    const je = context.journalEntry as Record<string, unknown>;
    parts.push(`\nJournal Entry Status:`);
    parts.push(`  - Total Debits: $${Number(je.totalDebit).toLocaleString()}`);
    parts.push(`  - Total Credits: $${Number(je.totalCredit).toLocaleString()}`);
    parts.push(`  - Balanced: ${je.isBalanced ? '✓ Yes' : '⚠️ No'}`);
  }

  if (context.dataQuality) {
    const dq = context.dataQuality as Record<string, unknown>;
    parts.push(`\nData Quality:`);
    if (dq.dataCleanReady) parts.push(`  - PR_Data_Clean: ✓ Ready`);
    if (dq.jeDraftReady) parts.push(`  - PR_JE_Draft: ✓ Ready`);
    if (dq.periodsAvailable) parts.push(`  - Historical Periods: ${dq.periodsAvailable}`);
  }

  return parts.join('\n') || 'No spreadsheet context available.';
}
