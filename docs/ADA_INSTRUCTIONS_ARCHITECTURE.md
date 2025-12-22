# Ada Instructions Architecture

## Current State (What You Already Have)

Your website already has a solid Ada admin foundation in place:

### Existing Tables (in `prairie-forge-website`)
1. **`ada_system_prompts`** - System prompts by name (e.g., `default`, `payroll-recorder`)
2. **`ada_conversations`** - Conversation logs with customer/module tracking

### Existing Admin UI
- **`AdaAdminPage.tsx`** - Full CRUD for prompts, modules, and conversation logs
- Three tabs: Modules, System Prompts, Conversations
- Module config with quick actions and context-specific instructions

### What's Missing
1. **`ada_module_config`** table - Referenced in UI but no migration exists
2. **`ada_knowledge_sources`** table - For FAQ/docs Ada should reference
3. **Integration** - The Payroll Recorder callAdaApi needs to pass module/function
4. **Hierarchical prompt resolution** - Currently flat name lookup

---

## Recommended Schema Enhancements

### 1. Create `ada_module_config` (Missing Table)

This is already referenced in your UI but needs the migration:

```sql
-- ============================================
-- ADA MODULE CONFIGURATION
-- ============================================
CREATE TABLE ada_module_config (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- Module identification
  module_key TEXT NOT NULL UNIQUE,              -- 'payroll-recorder', 'pto-accrual'
  display_name TEXT NOT NULL,                   -- 'Payroll Recorder'
  description TEXT,
  
  -- Link to base system prompt
  prompt_name TEXT REFERENCES ada_system_prompts(name),
  
  -- Context-specific instructions (appended to base prompt)
  ada_context_mapping TEXT,                     -- When helping with column mapping
  ada_context_analysis TEXT,                    -- When analyzing data
  ada_context_validation TEXT,                  -- When validating before export
  
  -- Quick action buttons in the UI
  ada_quick_actions JSONB DEFAULT '[]',
  
  -- Status
  is_active BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Seed data
INSERT INTO ada_module_config (module_key, display_name, description, prompt_name, ada_quick_actions)
VALUES
  ('payroll-recorder', 'Payroll Recorder', 'Payroll expense analysis and journal entry', 'payroll-recorder', 
   '[{"id": "diagnose", "label": "Run Diagnostics", "prompt": "Run a diagnostic check on the current payroll data"},
     {"id": "executive", "label": "Executive Summary", "prompt": "Generate an executive-ready summary of this payroll period"},
     {"id": "variance", "label": "Explain Variances", "prompt": "Identify and explain any significant variances from prior period"}]'),
  ('pto-accrual', 'PTO Accrual', 'PTO liability tracking and validation', 'default',
   '[{"id": "review", "label": "Review Accruals", "prompt": "Review the current PTO accrual calculations for accuracy"}]');
```

### 2. Create `ada_knowledge_sources` (New Table)

For FAQ, docs, and policies Ada should reference:

```sql
-- ============================================
-- ADA KNOWLEDGE SOURCES
-- ============================================
CREATE TABLE ada_knowledge_sources (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- Scope (NULL = applies globally)
  module_key TEXT,                              -- NULL or 'payroll-recorder', 'pto-accrual'
  function_context TEXT,                        -- NULL or 'mapping', 'analysis', 'validation'
  
  -- Content classification
  source_type TEXT NOT NULL CHECK (source_type IN ('faq', 'documentation', 'policy', 'workflow', 'glossary')),
  
  -- The actual knowledge
  title TEXT NOT NULL,
  content TEXT NOT NULL,
  keywords TEXT[],                              -- For matching user questions
  
  -- Ordering (higher = checked first)
  priority INT DEFAULT 50,
  
  -- Status
  is_active BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_ada_knowledge_module ON ada_knowledge_sources(module_key);
CREATE INDEX idx_ada_knowledge_type ON ada_knowledge_sources(source_type);
CREATE INDEX idx_ada_knowledge_active ON ada_knowledge_sources(is_active);

-- Example seed data
INSERT INTO ada_knowledge_sources (module_key, source_type, title, content, keywords, priority)
VALUES
  ('payroll-recorder', 'faq', 'Why is my journal entry not balancing?',
   'Journal entries must have equal debits and credits. Common causes include:
   • Missing GL account mappings for some expense categories
   • Rounding differences in the source data
   • Excluded rows that still affect totals
   
   To fix: Go to the Column Mapping step and ensure all categories have GL accounts assigned.',
   ARRAY['balance', 'debit', 'credit', 'not balancing', 'journal entry'],
   100),
   
  ('payroll-recorder', 'glossary', 'Burden Rate',
   'The burden rate represents the ratio of employer-paid costs (taxes, benefits, insurance) to gross wages.
   • Typical range: 15-25%
   • High burden (>30%): May indicate benefits-heavy workforce or calculation error
   • Low burden (<12%): May indicate missing employer tax categories',
   ARRAY['burden', 'burden rate', 'employer costs'],
   50),
   
  (NULL, 'workflow', 'How to Export to QuickBooks',
   'To export your journal entry to QuickBooks Online:
   1. Complete all mapping steps with green checkmarks
   2. Review the Journal Summary tab
   3. Click "Export to QBO" button
   4. Sign in to QuickBooks if prompted
   5. Confirm the import in QuickBooks',
   ARRAY['export', 'quickbooks', 'qbo', 'journal entry'],
   80);

-- RLS
ALTER TABLE ada_knowledge_sources ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Admins can manage knowledge" ON ada_knowledge_sources
  FOR ALL USING (
    EXISTS (
      SELECT 1 FROM profiles 
      WHERE profiles.id = auth.uid() 
      AND profiles.is_admin = true
    )
  );

CREATE POLICY "Anyone can read active knowledge" ON ada_knowledge_sources
  FOR SELECT USING (is_active = true);

GRANT SELECT ON ada_knowledge_sources TO anon;
GRANT ALL ON ada_knowledge_sources TO service_role;
```

---

## Enhanced Copilot Edge Function

Update the copilot function to:
1. Accept `module` and `function` parameters
2. Fetch relevant knowledge sources
3. Build composite prompt

```typescript
// In supabase/functions/copilot/index.ts

interface AdaRequest {
  prompt: string;
  context?: Record<string, unknown>;
  systemPrompt?: string;           // Optional override
  module?: string;                 // 'payroll-recorder', 'pto-accrual'
  function?: string;               // 'mapping', 'analysis', 'validation'
  promptName?: string;             // Direct lookup by name
  history?: Array<{ role: string; content: string }>;
  sessionId?: string;
  customerId?: string;
}

// Fetch module config and build composite prompt
async function buildCompositePrompt(
  module: string,
  func: string,
  baseSystemPrompt: string | null
): Promise<string> {
  const supabase = getSupabaseClient();
  if (!supabase) return baseSystemPrompt || DEFAULT_SYSTEM_PROMPT;
  
  // 1. Get module config
  const { data: moduleConfig } = await supabase
    .from('ada_module_config')
    .select('prompt_name, ada_context_mapping, ada_context_analysis, ada_context_validation')
    .eq('module_key', module)
    .eq('is_active', true)
    .single();
  
  // 2. Get base system prompt
  let prompt = baseSystemPrompt;
  if (!prompt && moduleConfig?.prompt_name) {
    const { data: promptData } = await supabase
      .from('ada_system_prompts')
      .select('prompt_text')
      .eq('name', moduleConfig.prompt_name)
      .eq('is_active', true)
      .single();
    prompt = promptData?.prompt_text;
  }
  prompt = prompt || DEFAULT_SYSTEM_PROMPT;
  
  // 3. Append function-specific context
  if (moduleConfig) {
    const contextMap: Record<string, string | null> = {
      'mapping': moduleConfig.ada_context_mapping,
      'analysis': moduleConfig.ada_context_analysis,
      'validation': moduleConfig.ada_context_validation,
    };
    const contextAddition = contextMap[func];
    if (contextAddition) {
      prompt += `\n\n## ADDITIONAL CONTEXT FOR ${func.toUpperCase()}\n${contextAddition}`;
    }
  }
  
  // 4. Inject relevant knowledge sources
  const { data: knowledge } = await supabase
    .from('ada_knowledge_sources')
    .select('source_type, title, content')
    .or(`module_key.is.null,module_key.eq.${module}`)
    .or(`function_context.is.null,function_context.eq.${func}`)
    .eq('is_active', true)
    .order('priority', { ascending: false })
    .limit(5);
  
  if (knowledge?.length) {
    prompt += '\n\n## REFERENCE KNOWLEDGE\n';
    prompt += 'Use this knowledge to answer user questions when relevant:\n\n';
    for (const k of knowledge) {
      prompt += `### [${k.source_type.toUpperCase()}] ${k.title}\n${k.content}\n\n`;
    }
  }
  
  return prompt;
}
```

---

## Client Integration (Payroll Recorder)

Update `callAdaApi()` to pass module/function:

```javascript
// In workflow.js
async function callAdaApi(params) {
    const { systemPrompt, userPrompt, contextPack, functionContext } = params;
    
    const response = await fetch(COPILOT_URL, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${SUPABASE_ANON_KEY}`,
            "apikey": SUPABASE_ANON_KEY
        },
        body: JSON.stringify({
            prompt: userPrompt,
            context: contextPack,
            module: "payroll-recorder",           // <-- NEW
            function: functionContext || "analysis",  // <-- NEW
            systemPrompt: systemPrompt            // Optional override
        })
    });
    // ...
}
```

---

## Admin Portal Enhancements

### Tab 1: Modules (Already Exists ✓)
Your existing UI is great. Just needs the migration to create `ada_module_config`.

### Tab 2: System Prompts (Already Exists ✓)
Works well. Consider adding:
- Version history
- "Duplicate" button for creating variations
- Preview with sample context

### Tab 3: Knowledge Base (NEW)
Add a new tab to manage FAQ/docs:

```tsx
// New tab in AdaAdminPage.tsx
<TabsTrigger value="knowledge" className="gap-2">
  <BookOpen className="h-4 w-4" />
  Knowledge Base
</TabsTrigger>

<TabsContent value="knowledge">
  {/* CRUD for ada_knowledge_sources */}
  {/* Fields: module, function, source_type, title, content, keywords, priority */}
  {/* Filter by module/type */}
</TabsContent>
```

### Tab 4: Conversations (Already Exists ✓)
Works well. Consider adding:
- Filter by module
- Cost tracking (tokens × $0.00003 for GPT-4 Turbo)
- "Unanswered" flag for questions Ada couldn't handle

---

## Implementation Order

### Phase 1: Database (Do First)
1. ✅ `ada_system_prompts` - Already exists
2. ✅ `ada_conversations` - Already exists  
3. ⏳ Create `ada_module_config` migration
4. ⏳ Create `ada_knowledge_sources` migration

### Phase 2: Copilot Edge Function
1. ⏳ Update to accept `module` and `function` params
2. ⏳ Add `buildCompositePrompt()` function
3. ⏳ Inject knowledge sources into prompt

### Phase 3: Client Integration
1. ⏳ Update Payroll Recorder `callAdaApi()` to pass module/function
2. ⏳ Update PTO Accrual to use same pattern

### Phase 4: Admin UI
1. ⏳ Create Knowledge Base tab
2. ⏳ Add analytics/cost tracking
3. ⏳ Add "test prompt" feature

---

## Next Steps

Would you like me to:

1. **Create the migrations** - Generate SQL files for `ada_module_config` and `ada_knowledge_sources`
2. **Update the copilot edge function** - Add hierarchical prompt resolution
3. **Update the Payroll Recorder** - Pass module/function to Ada
4. **Build the Knowledge Base UI** - Add tab to your existing AdaAdminPage

Which should I start with?
