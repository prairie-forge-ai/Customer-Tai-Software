# 📋 pf_column_name as Target — Example Responses
**Date:** December 11, 2025

---

## Confirmation

| Item | Status |
|------|--------|
| Amount dictionary lookups return `pf_column_name` | ✅ Confirmed |
| Dimension lookups return `normalized_dimension` | ✅ Confirmed |
| `normalized_key` is NOT exposed in response | ✅ Confirmed |
| Customer mappings save `normalized_key = pf_column_name` | ✅ Confirmed |

---

## 1. Saved Mapping (Company-Specific)

**Scenario:** Header "Regular Earns" has a saved mapping for this company.

```json
{
  "raw_header": "Regular Earns",
  "kind": "amount",
  "target": "Wages_Salary_Amount",
  "target_key": "Wages_Salary_Amount",
  "source": "saved",
  "confidence": 1.0,
  "gl_account": "61811",
  "gl_account_name": "Salaries & Wages"
}
```

**Notes:**
- `source: "saved"` indicates this came from `ada_customer_column_mappings`
- `target` = the `normalized_key` value stored (which equals `pf_column_name`)
- GL account enriched from `ada_customer_gl_mappings`

---

## 2. Amount Match (Dictionary — Exact)

**Scenario:** Header "Regular Pay" matches `data_source_name` in `ada_payroll_column_dictionary`.

```json
{
  "raw_header": "Regular Pay",
  "kind": "amount",
  "target": "Wages_Salary_Amount",
  "target_key": "Wages_Salary_Amount",
  "source": "amount",
  "confidence": 0.95,
  "gl_account": "61811",
  "gl_account_name": "Salaries & Wages"
}
```

**Notes:**
- `source: "amount"` indicates exact match from amount dictionary
- `target` = `pf_column_name` from `ada_payroll_column_dictionary`
- High confidence (0.95) for exact matches

---

## 3. Dimension Match (Exact)

**Scenario:** Header "Employee Name" matches `raw_term` in `ada_payroll_dimensions`.

```json
{
  "raw_header": "Employee Name",
  "kind": "dimension",
  "target": "Employee_Name",
  "target_key": "Employee_Name",
  "source": "dimension",
  "confidence": 0.95
}
```

**Notes:**
- `source: "dimension"` indicates exact match from dimension dictionary
- `target` = `normalized_dimension` from `ada_payroll_dimensions`
- No GL account (dimensions don't have GL mappings)

---

## 4. Ambiguous Case

**Scenario:** Header "Pay Date" exists in BOTH amount dictionary AND dimension dictionary.

```json
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
```

**Notes:**
- `kind: "ambiguous"` — user must choose
- `target: null` — no automatic selection
- `amount_option.target` = `pf_column_name` from amount dictionary
- `dimension_option.target` = `normalized_dimension` from dimension dictionary
- Low confidence (0.5) due to ambiguity

---

## 5. Fuzzy Match (Amount)

**Scenario:** Header "Reg Pay" partially matches "Regular Pay" in amount dictionary.

```json
{
  "raw_header": "Reg Pay",
  "kind": "amount",
  "target": "Wages_Salary_Amount",
  "target_key": "Wages_Salary_Amount",
  "source": "fuzzy",
  "confidence": 0.72
}
```

**Notes:**
- `source: "fuzzy"` indicates partial/fuzzy match
- `target` = `pf_column_name` (same as exact match)
- Lower confidence (< 0.8) for fuzzy matches

---

## 6. Fuzzy Match (Dimension)

**Scenario:** Header "Emp Name" partially matches "Employee Name" in dimension dictionary.

```json
{
  "raw_header": "Emp Name",
  "kind": "dimension",
  "target": "Employee_Name",
  "target_key": "Employee_Name",
  "source": "fuzzy",
  "confidence": 0.56
}
```

**Notes:**
- `source: "fuzzy"` indicates partial/fuzzy match
- Dimension fuzzy matches have even lower confidence (× 0.7)

---

## 7. Unmapped

**Scenario:** Header "XYZ Custom Field" has no match anywhere.

```json
{
  "raw_header": "XYZ Custom Field",
  "kind": null,
  "target": null,
  "target_key": null,
  "source": "unmapped",
  "confidence": 0
}
```

**Notes:**
- `kind: null` — no mapping type determined
- `target: null` — no target found
- `source: "unmapped"` — explicit unmapped status
- User must manually select a mapping or skip

---

## Full Response Example

**Request:**
```json
{
  "action": "analyze",
  "headers": [
    "Regular Earns",
    "Regular Pay",
    "Employee Name",
    "Pay Date",
    "Reg Pay",
    "Emp Name",
    "XYZ Custom Field"
  ],
  "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
  "module": "payroll-recorder"
}
```

**Response:**
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
      "raw_header": "Regular Pay",
      "kind": "amount",
      "target": "Wages_Salary_Amount",
      "target_key": "Wages_Salary_Amount",
      "source": "amount",
      "confidence": 0.95,
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
    },
    {
      "raw_header": "Reg Pay",
      "kind": "amount",
      "target": "Wages_Salary_Amount",
      "target_key": "Wages_Salary_Amount",
      "source": "fuzzy",
      "confidence": 0.72
    },
    {
      "raw_header": "Emp Name",
      "kind": "dimension",
      "target": "Employee_Name",
      "target_key": "Employee_Name",
      "source": "fuzzy",
      "confidence": 0.56
    },
    {
      "raw_header": "XYZ Custom Field",
      "kind": null,
      "target": null,
      "target_key": null,
      "source": "unmapped",
      "confidence": 0
    }
  ],
  "source": "saved",
  "matched": 4,
  "unmapped": 1,
  "amounts": 3,
  "dimensions": 2,
  "ambiguous": 1,
  "fuzzy": 2,
  "with_gl": 3,
  "total": 7
}
```

---

## Save Operation

When saving a mapping, `target` (which is `pf_column_name`) is stored as `normalized_key`:

**Request:**
```json
{
  "action": "save",
  "company_id": "d50495c3-79a4-4895-85f7-6b0bf6c409d8",
  "module": "payroll-recorder",
  "mappings": [
    {
      "raw_header": "Regular Earns",
      "target": "Wages_Salary_Amount",
      "kind": "amount"
    },
    {
      "raw_header": "Employee Name",
      "target": "Employee_Name",
      "kind": "dimension"
    }
  ]
}
```

**Database INSERT:**
```sql
INSERT INTO ada_customer_column_mappings 
  (company_id, module, raw_header, normalized_key, mapping_type, confidence, source)
VALUES 
  ('d50495c3-...', 'payroll-recorder', 'Regular Earns', 'Wages_Salary_Amount', 'amount', 1.0, 'ada_confirmed'),
  ('d50495c3-...', 'payroll-recorder', 'Employee Name', 'Employee_Name', 'dimension', 1.0, 'ada_confirmed');
```

**Response:**
```json
{
  "success": true,
  "saved": 2,
  "errors": []
}
```

---

## Summary Table

| Source | `target` comes from |
|--------|---------------------|
| `saved` | `ada_customer_column_mappings.normalized_key` (= `pf_column_name`) |
| `amount` | `ada_payroll_column_dictionary.pf_column_name` |
| `dimension` | `ada_payroll_dimensions.normalized_dimension` |
| `fuzzy` | Same as above (amount or dimension) |
| `ambiguous` | `null` — options provided separately |
| `unmapped` | `null` |












