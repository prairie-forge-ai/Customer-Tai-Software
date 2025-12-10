/**
 * Ada - Prairie Forge AI Assistant
 * Supabase Edge Function powered by ChatGPT
 * 
 * Features:
 * - Fetches system prompts from database (admin-editable)
 * - Logs all conversations for debugging and improvement
 * - Supports multiple prompt personalities
 * 
 * COST ESTIMATES (GPT-4 Turbo):
 * - Typical question: ~$0.02-0.05
 * - 100 questions/day ≈ $2-5/day
 */

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// Configuration
const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") || "https://jgciqwzwacaesqjaoadc.supabase.co";
const SUPABASE_SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

// Default model configuration (can be overridden by database)
const DEFAULT_MODEL = "gpt-4-turbo-preview";
const DEFAULT_MAX_TOKENS = 1500;
const DEFAULT_TEMPERATURE = 0.7;

// CORS headers for Excel add-in
const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

interface AdaRequest {
  prompt: string;
  context?: Record<string, unknown>;
  systemPrompt?: string;
  promptName?: string; // Name of the system prompt to use from database
  history?: Array<{ role: string; content: string }>;
  sessionId?: string;
  customerId?: string;
}

interface SystemPromptConfig {
  prompt_text: string;
  model: string;
  max_tokens: number;
  temperature: number;
}

// Create Supabase client for database operations
function getSupabaseClient() {
  if (!SUPABASE_SERVICE_KEY) {
    console.warn("SUPABASE_SERVICE_ROLE_KEY not set, database features disabled");
    return null;
  }
  return createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
}

// Fetch system prompt from database
async function fetchSystemPrompt(promptName: string = "default"): Promise<SystemPromptConfig | null> {
  const supabase = getSupabaseClient();
  if (!supabase) return null;

  try {
    const { data, error } = await supabase
      .from('ada_system_prompts')
      .select('prompt_text, model, max_tokens, temperature')
      .eq('name', promptName)
      .eq('is_active', true)
      .single();

    if (error) {
      console.error('Error fetching system prompt:', error);
      return null;
    }

    return data;
  } catch (e) {
    console.error('Failed to fetch system prompt:', e);
    return null;
  }
}

// Log conversation to database
async function logConversation(
  request: AdaRequest,
  response: string | null,
  tokensUsed: number | null,
  latencyMs: number,
  error: string | null,
  model: string
) {
  const supabase = getSupabaseClient();
  if (!supabase) return;

  try {
    await supabase
      .from('ada_conversations')
      .insert({
        session_id: request.sessionId || null,
        customer_id: request.customerId || null,
        prompt_name: request.promptName || 'default',
        user_prompt: request.prompt,
        context: request.context || null,
        ai_response: response,
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
  
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  let body: AdaRequest | null = null;
  let model = DEFAULT_MODEL;

  try {
    // Validate API key is configured
    if (!OPENAI_API_KEY) {
      console.error("OPENAI_API_KEY not configured");
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
    const { prompt, context, systemPrompt, promptName, history } = body;

    if (!prompt?.trim()) {
      return new Response(
        JSON.stringify({ error: "Please ask Ada a question!", code: "INVALID_REQUEST" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Try to fetch system prompt from database
    let promptConfig: SystemPromptConfig | null = null;
    let effectiveSystemPrompt = systemPrompt;
    let maxTokens = DEFAULT_MAX_TOKENS;
    let temperature = DEFAULT_TEMPERATURE;

    if (!systemPrompt) {
      promptConfig = await fetchSystemPrompt(promptName || 'default');
      if (promptConfig) {
        effectiveSystemPrompt = promptConfig.prompt_text;
        model = promptConfig.model || DEFAULT_MODEL;
        maxTokens = promptConfig.max_tokens || DEFAULT_MAX_TOKENS;
        temperature = promptConfig.temperature || DEFAULT_TEMPERATURE;
      }
    }

    // Build messages for OpenAI
    const messages = buildMessages(prompt, context, effectiveSystemPrompt, history);

    // Call OpenAI
    const openaiResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        messages,
        max_tokens: maxTokens,
        temperature,
      }),
    });

    if (!openaiResponse.ok) {
      const errorData = await openaiResponse.json();
      console.error("OpenAI API error:", errorData);
      
      const latencyMs = Date.now() - startTime;
      await logConversation(body, null, null, latencyMs, `OpenAI error: ${openaiResponse.status}`, model);
      
      if (openaiResponse.status === 429) {
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

    const completion = await openaiResponse.json();
    const responseMessage = completion.choices?.[0]?.message?.content || "I couldn't generate a response. Please try rephrasing your question.";
    const tokensUsed = completion.usage?.total_tokens || 0;
    const latencyMs = Date.now() - startTime;

    // Log successful conversation
    await logConversation(body, responseMessage, tokensUsed, latencyMs, null, model);

    console.log(`Ada responded: ${tokensUsed} tokens used, ${latencyMs}ms`);

    // Return successful response
    return new Response(
      JSON.stringify({
        message: responseMessage,
        usage: {
          tokens: tokensUsed,
          model: model,
          latencyMs: latencyMs
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
      await logConversation(body, null, null, latencyMs, String(error), model);
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
 * Build the messages array for OpenAI
 */
function buildMessages(
  prompt: string, 
  context: Record<string, unknown> | undefined,
  systemPrompt: string | undefined,
  history: Array<{ role: string; content: string }> | undefined
): Array<{ role: string; content: string }> {
  const messages: Array<{ role: string; content: string }> = [];

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

  messages.push({
    role: "system",
    content: systemPrompt || defaultSystemPrompt
  });

  // Add context if provided
  if (context && Object.keys(context).length > 0) {
    const contextSummary = formatContextForAI(context);
    messages.push({
      role: "system",
      content: `Current spreadsheet data:\n${contextSummary}`
    });
  }

  // Add conversation history (last 8 messages for good context)
  if (history?.length) {
    const recentHistory = history.slice(-8);
    for (const msg of recentHistory) {
      if (msg.role === "user" || msg.role === "assistant") {
        messages.push({ role: msg.role, content: msg.content });
      }
    }
  }

  // Add current prompt
  messages.push({ role: "user", content: prompt });

  return messages;
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
