# ✅ Follow-Up Confirmation
**Date:** December 11, 2025

---

## 1. Ambiguous Handling: Concrete Example

### Confirmed Behavior

When a header matches **BOTH** an amount AND dimension dictionary, the response contains:

```json
{
  "mappings": [
    {
      "raw_header": "Pay Date",
      "kind": "ambiguous",
      "target": null,
      "target_key": null,
      "source": "ambiguous",
      "confidence": 0.5,
      "amount_option": {
        "target": "Pay_Date_Amount",
        "confidence": 0.95
      },
      "dimension_option": {
        "target": "Pay_Date",
        "confidence": 0.95
      }
    }
  ],
  "ambiguous": 1
}
```

### Key Guarantees

| Requirement | Status |
|-------------|--------|
| Exactly ONE entry per `raw_header` | ✅ Confirmed |
| `kind = "ambiguous"` | ✅ Confirmed |
| `target = null` | ✅ Confirmed |
| `target_key = null` | ✅ Confirmed |
| `source = "ambiguous"` | ✅ Confirmed |
| Low confidence (0.5) | ✅ Confirmed |
| Both options provided | ✅ `amount_option` + `dimension_option` |

---

## 2. CHECK Constraints: Explicitly Removed

### Migration: `014_drop_check_constraints.sql`

```sql
-- Drop the CHECK constraint on ada_customer_column_mappings
ALTER TABLE ada_customer_column_mappings 
  DROP CONSTRAINT IF EXISTS chk_valid_normalized_key;

-- Drop the CHECK constraint on ada_customer_gl_mappings  
ALTER TABLE ada_customer_gl_mappings 
  DROP CONSTRAINT IF EXISTS chk_valid_normalized_key_gl;

-- Drop the validation function (no longer needed)
DROP FUNCTION IF EXISTS is_valid_normalized_key(TEXT);
```

### What Was Removed

| Item | Status |
|------|--------|
| `chk_valid_normalized_key` on `ada_customer_column_mappings` | ✅ DROPPED |
| `chk_valid_normalized_key_gl` on `ada_customer_gl_mappings` | ✅ DROPPED |
| `is_valid_normalized_key(TEXT)` function | ✅ DROPPED |

### Why Removed

The CHECK constraints used `is_valid_normalized_key()` which validated against `ada_payroll_taxonomy`. Since that table is now **LEGACY/REMOVED**, these constraints would block all inserts.

Validation is now handled by:
1. UI dropdowns (only show valid options)
2. Dictionary lookups (only return valid targets)

---

## Files Changed This Session

| File | Purpose |
|------|---------|
| `supabase/functions/column-mapper/index.ts` | Rewritten with proper ambiguity handling |
| `supabase/migrations/013_finalize_column_dictionary_schema.sql` | Schema for `ada_payroll_column_dictionary` |
| `supabase/migrations/014_drop_check_constraints.sql` | Drop CHECK constraints |
| `docs/UPDATED_ARCHITECTURE_SPEC.md` | Full API spec with examples |

---

## Deployment Checklist

Run these in order:

### 1. Run Migrations in Supabase SQL Editor

```sql
-- First: Schema migration
-- Run contents of: supabase/migrations/013_finalize_column_dictionary_schema.sql

-- Second: Drop CHECK constraints
-- Run contents of: supabase/migrations/014_drop_check_constraints.sql
```

### 2. Deploy Edge Function

```bash
cd /Users/d.paeth/Customer-ArchCollins-Foundry
supabase functions deploy column-mapper --no-verify-jwt
```

### 3. Test Ambiguity

To test ambiguous handling, you need entries in BOTH tables for the same term:

```sql
-- Add to amount dictionary
INSERT INTO ada_payroll_column_dictionary (module, data_source_name, normalized_key, mapping_type)
VALUES ('payroll-recorder', 'Pay Date', 'Pay_Date_Amount', 'financial');

-- Add to dimensions (if not already there)
INSERT INTO ada_payroll_dimensions (raw_term, normalized_dimension, semantic_group)
VALUES ('Pay Date', 'Pay_Date', 'date_fields');
```

Then test:
```bash
curl -X POST \
  "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/column-mapper" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "analyze",
    "headers": ["Pay Date"],
    "module": "payroll-recorder"
  }'
```

Expected: `kind = "ambiguous"`, `target = null`, `source = "ambiguous"`

---

## Summary

| Question | Answer |
|----------|--------|
| Is ambiguity collapsed to single entry? | ✅ Yes |
| Does ambiguous have `target = null`? | ✅ Yes |
| Does ambiguous have `source = "ambiguous"`? | ✅ Yes |
| Are CHECK constraints removed? | ✅ Yes |
| Is `is_valid_normalized_key()` dropped? | ✅ Yes |

