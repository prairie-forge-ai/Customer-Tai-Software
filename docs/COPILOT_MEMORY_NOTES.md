# Prairie Forge AI - Memory Notes for Future Chats

## Project Overview
Prairie Forge builds Excel add-ins for accounting workflows. Two main repos:
- **Customer-ArchCollins-Foundry**: Excel add-in modules (payroll-recorder, pto-accrual), Supabase edge functions
- **prairie-forge-website** (`~/prairie-forge-website`): React admin portal, shared database migrations

Supabase Project ID: `jgciqwzwacaesqjaoadc`

---

## Ada AI Assistant Architecture

Ada is the AI assistant embedded in Excel add-ins, powered by GPT-4 Turbo.

### Key Tables
- `ada_system_prompts` - Base system prompts with model settings (temperature, max_tokens)
- `ada_module_config` - Per-module configuration (prompt_name, ada_context_mapping/analysis/validation, quick_actions)
- `ada_knowledge_sources` - FAQs, docs, policies Ada references (source_type, keywords, priority)
- `ada_conversations` - Conversation history with token usage and latency metrics

### Edge Functions
- `copilot` - Main Ada AI endpoint, calls OpenAI GPT-4 Turbo
- `column-mapper` - AI-powered column mapping for payroll data imports

### Prompt Resolution Flow
1. Client passes `module` and `function` parameters
2. copilot fetches `ada_module_config` for the module
3. Gets base prompt from `ada_system_prompts` (via prompt_name)
4. Appends function-specific context (ada_context_mapping/analysis/validation)
5. Injects relevant knowledge from `ada_knowledge_sources`

### Client-Side Pattern
```javascript
await fetch('https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/copilot', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
  },
  body: JSON.stringify({
    prompt: userQuestion,
    context: contextData,
    module: 'payroll-recorder',  // or 'pto-accrual'
    function: 'analysis',  // or 'mapping', 'validation'
    sessionId: sessionId,
    customerId: customerId,
  }),
});
```

---

## Build & Deploy Commands

### Payroll Recorder
```bash
cd ~/Customer-ArchCollins-Foundry
node scripts/build-payroll.js
```

### Deploy Edge Function
```bash
npx supabase functions deploy copilot
npx supabase functions deploy column-mapper
```

### Supabase Secrets
```bash
npx supabase secrets list
npx supabase secrets set OPENAI_API_KEY=sk-...
```

---

## Database Patterns

### RLS Policy Pattern
User's Supabase uses `user_roles` table for admin checks, NOT `profiles.is_admin`:
```sql
EXISTS (
  SELECT 1 FROM user_roles 
  WHERE user_roles.user_id = auth.uid() 
  AND user_roles.role = 'admin'
)
```

### Knowledge Sources Structure
- `source_type`: faq | documentation | policy | workflow | glossary | troubleshooting
- `module_key`: NULL (global) or specific module key
- `function_context`: NULL (all functions) or mapping/analysis/validation
- `keywords`: text[] for matching user questions
- `priority`: higher = checked first (default 50)

---

## Admin Portal Patterns

### ModuleInfoButton Component
All admin pages should include technical details button:
```tsx
import { ModuleInfoButton } from "@/components/admin/ModuleInfoButton";

<ModuleInfoButton
  moduleName="Ada AI Assistant"
  description="AI-powered assistant for Excel add-ins"
  tables={[
    { name: 'ada_system_prompts', description: 'Base system prompts' },
    { name: 'ada_module_config', description: 'Per-module configuration' },
  ]}
  functions={[
    { name: 'copilot', description: 'Edge Function: Main Ada AI endpoint' },
  ]}
  storageBuckets={[]}
  views={[]}
/>
```

---

## Code Conventions

### Module Keys
- `payroll-recorder` - Payroll expense recording module
- `pto-accrual` - PTO accrual tracking module

### Sheet Naming
- `SS_*` prefix - System sheets (SS_PF_Config, SS_Employee_Roster)
- `PR_*` prefix - Payroll Recorder sheets (PR_Data_Clean, PR_JE_Draft)
- `PTO_*` prefix - PTO Accrual sheets

### Config Field Naming
Pattern: `{MODULE}_{Descriptor}`
- `PR_Payroll_Date`, `PR_Accounting_Period`, `PR_JE_Debit_Total`
- `PTO_Accounting_Period`, `PTO_Accrual_Rate`

---

## TypeScript Notes

### Supabase Edge Functions
- Use Deno runtime, NOT Node.js
- VS Code shows TypeScript errors for `Deno.env.get()` - these are IDE-only errors
- Code works fine when deployed
- Fix: Install Deno VS Code extension or add tsconfig to exclude

---

## Migration Files
Once migrations are executed in Supabase SQL Editor, the .sql files are historical records.
Safe to delete, but many teams keep them for documentation.

---

## Key Documentation Files
- `/docs/ADA_INSTRUCTIONS_ARCHITECTURE.md` - Full Ada system architecture
- `/payroll-recorder/README-payroll-recorder.md` - Payroll module docs
- `/pto-accrual/README-pto-accrual.md` - PTO module docs
