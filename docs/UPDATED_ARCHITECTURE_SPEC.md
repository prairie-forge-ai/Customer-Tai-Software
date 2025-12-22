# 📋 Updated Column Mapper Architecture
**Date:** December 11, 2025  
**Status:** Implementation Complete (pending deployment)

---

## Source of Truth Tables

| Table | Purpose | Key Columns |
|-------|---------|-------------|
| `ada_payroll_column_dictionary` | **AMOUNTS** (financial columns) | `module`, `provider`, `data_source_name`, `normalized_key` |
| `ada_payroll_dimensions` | **DIMENSIONS** (identity/metadata) | `provider`, `raw_term`, `normalized_dimension`, `semantic_group` |
| `ada_customer_column_mappings` | **SAVED** (company-specific) | `company_id`, `module`, `raw_header`, `normalized_key`, `mapping_type` |
| `ada_customer_gl_mappings` | **GL LOOKUP** (account numbers) | `company_id`, `module`, `normalized_key`, `gl_account` |

**⚠️ `ada_payroll_taxonomy` is REMOVED/LEGACY — do not use.**

---

## Lookup Order

```
1. ada_customer_column_mappings  → source: "saved"      (confidence: 1.0)
2. ada_payroll_column_dictionary → source: "amount"     (confidence: 0.95)  [exact match]
3. ada_payroll_dimensions        → source: "dimension"  (confidence: 0.95)  [exact match]
4. ada_payroll_column_dictionary → source: "fuzzy"      (confidence: 0.5-0.8) [fuzzy]
5. ada_payroll_dimensions        → source: "fuzzy"      (confidence: 0.4-0.7) [fuzzy]
6. No match                      → source: "unmapped"   (confidence: 0)
```

---

## API Contract

### Analyze Request

```json
POST /column-mapper
{
  "action": "analyze",
  "headers": ["Regular Earns", "Employee Name", "Department", "Unknown"],
  "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
  "module": "payroll-recorder"
}
```

### Analyze Response

```json
{
  "mappings": [
    {
      "raw_header": "Regular Earns",
      "kind": "amount",
      "target": "Wages_Salary_Amount",
      "target_key": "Wages_Salary_Amount",
      "source": "saved",
      "confidence": 1.0,
      "gl_account": "61811",
      "gl_account_name": "Salaries & Wages"
    },
    {
      "raw_header": "Employee Name",
      "kind": "dimension",
      "target": "Employee_Name",
      "target_key": "Employee_Name",
      "source": "dimension",
      "confidence": 0.95
    },
    {
      "raw_header": "Department",
      "kind": "dimension",
      "target": "Department",
      "target_key": "Department",
      "source": "dimension",
      "confidence": 0.95
    },
    {
      "raw_header": "Unknown",
      "kind": null,
      "target": null,
      "target_key": null,
      "source": "unmapped",
      "confidence": 0
    }
  ],
  "source": "saved",
  "matched": 3,
  "unmapped": 1,
  "amounts": 1,
  "dimensions": 2,
  "ambiguous": 0,
  "with_gl": 1,
  "total": 4
}
```

### Save Request

```json
POST /column-mapper
{
  "action": "save",
  "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
  "module": "payroll-recorder",
  "mappings": [
    { "raw_header": "Regular Earns", "target": "Wages_Salary_Amount", "kind": "amount" },
    { "raw_header": "Employee Name", "target": "Employee_Name", "kind": "dimension" }
  ]
}
```

### Save Response

```json
{
  "success": true,
  "saved": 2,
  "errors": []
}
```

---

## Response Field Definitions

| Field | Type | Description |
|-------|------|-------------|
| `raw_header` | string | Original column header from uploaded file |
| `kind` | `"amount"` \| `"dimension"` \| `"ambiguous"` \| `null` | Type of mapping |
| `target` | string \| null | PF canonical field name (e.g., `Wages_Salary_Amount`) |
| `target_key` | string \| null | Same as `target` for now |
| `source` | `"saved"` \| `"amount"` \| `"dimension"` \| `"fuzzy"` \| `"unmapped"` | Where the mapping came from |
| `confidence` | number | 0.0-1.0, higher = more certain |
| `gl_account` | string \| null | GL account code (amounts only) |
| `gl_account_name` | string \| null | GL account description |

---

## `ada_payroll_column_dictionary` Schema

```sql
CREATE TABLE ada_payroll_column_dictionary (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    module TEXT NOT NULL DEFAULT 'payroll-recorder',
    provider TEXT DEFAULT 'Global',
    data_source_name TEXT NOT NULL,     -- Header as it appears in source
    normalized_key TEXT NOT NULL,        -- PF canonical: "Wages_Salary_Amount"
    mapping_type TEXT DEFAULT 'financial',
    notes TEXT,
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now()
);

-- Indexes
CREATE INDEX idx_apcd_module ON ada_payroll_column_dictionary(module);
CREATE INDEX idx_apcd_data_source_name ON ada_payroll_column_dictionary(LOWER(data_source_name));
CREATE UNIQUE INDEX idx_apcd_unique_mapping ON ada_payroll_column_dictionary(module, provider, LOWER(data_source_name));
```

---

## `ada_payroll_dimensions` Schema

```sql
CREATE TABLE ada_payroll_dimensions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    provider TEXT DEFAULT 'Global',
    raw_term TEXT NOT NULL,              -- Header as it appears in source
    normalized_dimension TEXT NOT NULL,   -- PF canonical: "Employee_Name"
    data_type TEXT,                       -- "text", "date", "number"
    semantic_group TEXT,                  -- "employee_identity", "org_structure"
    is_alias BOOLEAN DEFAULT false,
    notes TEXT,
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now()
);
```

---

## `ada_customer_column_mappings` Schema

```sql
CREATE TABLE ada_customer_column_mappings (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL,
    module TEXT NOT NULL,
    raw_header TEXT NOT NULL,
    normalized_key TEXT NOT NULL,         -- Maps to target from dictionary/dimensions
    mapping_type TEXT DEFAULT 'amount',   -- "amount" or "dimension"
    confidence FLOAT DEFAULT 1.0,
    source TEXT,                          -- "ada_confirmed", "user_manual"
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now(),
    
    UNIQUE(company_id, module, raw_header)
);
```

---

## Ambiguity Handling

If a header matches BOTH an amount dictionary entry AND a dimension dictionary entry,
**exactly ONE entry** is returned with:

- `kind = "ambiguous"`
- `target = null`
- `target_key = null`
- `source = "ambiguous"`
- `confidence = 0.5` (low confidence)
- `amount_option` = the amount match details
- `dimension_option` = the dimension match details

### Concrete Example

**Scenario:** Header "Pay Date" exists in BOTH:
- `ada_payroll_column_dictionary` as `data_source_name = "Pay Date"` → `normalized_key = "Pay_Date_Amount"`
- `ada_payroll_dimensions` as `raw_term = "Pay Date"` → `normalized_dimension = "Pay_Date"`

**Request:**
```json
{
  "action": "analyze",
  "headers": ["Pay Date"],
  "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
  "module": "payroll-recorder"
}
```

**Response:**
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
  "source": "dictionary",
  "matched": 0,
  "unmapped": 0,
  "amounts": 0,
  "dimensions": 0,
  "ambiguous": 1,
  "with_gl": 0,
  "total": 1
}
```

**Key points:**
- Only ONE entry in `mappings` for "Pay Date"
- `kind = "ambiguous"` (not "amount" or "dimension")
- `target = null` (user must choose)
- `source = "ambiguous"` (explicitly marked)
- Both options provided for UI to display

---

## Files Changed

| File | Change |
|------|--------|
| `supabase/functions/column-mapper/index.ts` | Complete rewrite with new architecture |
| `supabase/migrations/013_finalize_column_dictionary_schema.sql` | Schema migration for dictionary table |

---

## Deployment Steps

1. Run migration in Supabase SQL Editor:
   - `013_finalize_column_dictionary_schema.sql`

2. Deploy Edge Function:
   ```bash
   cd /path/to/repo
   supabase functions deploy column-mapper --no-verify-jwt
   ```

3. Populate `ada_payroll_column_dictionary` with entries (user to provide data)

4. Test with curl:
   ```bash
   curl -X POST \
     "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/column-mapper" \
     -H "Content-Type: application/json" \
     -d '{
       "action": "analyze",
       "headers": ["Regular Pay", "Employee Name"],
       "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
       "module": "payroll-recorder"
     }'
   ```

---

## Next Steps

1. **Populate `ada_payroll_column_dictionary`** with amount entries for `payroll-recorder` module
2. **Verify `ada_payroll_dimensions`** has dimension entries
3. **Deploy Edge Function** to Supabase
4. **Test end-to-end** in Excel add-in

