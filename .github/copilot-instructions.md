# Copilot Instructions for Prairie Forge Modules

## Project Overview
Excel/PowerPoint add-ins (task panes) using ES modules bundled with esbuild. Supabase backend with PostgreSQL + Edge Functions (Deno).

## Architecture Rules
- **Bundling**: Always run `npm run build:payroll` (or `:pto`, `:all`) after editing module source files
- **Shared code**: `Common/` utilities are imported into bundles; changes require rebuilding dependent modules
- **Entry points**: Each module has `src/` (edit this), `app.bundle.js` (generated - never edit)

## Code Patterns
- Use ES module syntax (`import`/`export`) - no global `window.PrairieForge` namespace
- Shared CSS in `Common/styles.css` - avoid module-level `:root` overrides
- Supabase Edge Functions use Deno runtime (TypeScript)
- RLS policies use `user_roles.role = 'admin'` (NOT `profiles.is_admin`)

## Key Commands
```bash
npm run build:payroll     # Bundle payroll-recorder module
npm run build:all         # Bundle all modules
npx supabase functions deploy copilot  # Deploy Ada AI function
```

## File Locations
- Module source: `<module>/src/*.js`
- Build scripts: `scripts/build-*.js`
- Edge functions: `supabase/functions/*/index.ts`
- Documentation: `docs/`

## Testing & Validation
- Run `npm test` for metrics tests
- Check bundle output exists after builds
- Verify edge function deployment with `supabase functions list`

## Important Context
- Ada AI uses GPT-4 Turbo via `/supabase/functions/copilot`
- Module configs stored in `ada_module_config` table
- Knowledge sources (FAQs) in `ada_knowledge_sources` table
- Project ID: `jgciqwzwacaesqjaoadc`
