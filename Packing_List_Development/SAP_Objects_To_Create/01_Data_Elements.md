# 01 — Data elements (SE11)

**No new data elements for the FS fields** — the existing S2A data elements below are reused
(as in the program version reviewed in the system).

| Field(s) | Data element (existing) |
|---|---|
| `ZSD_PACKNO` | `ZSD_VBELN` |
| `ZSD_SN`, `ZSD_INV_SN` | `ZSERNO` |
| `ZSD_ARTNO` | `ZCHAR15` |
| `ZSD_STD`, `ZSD_ART`, `ZSD_TART` | `ZSQNO` |
| `ZSD_PARTNO`, `ZSD_MATDESC` | `ZSD_DEBEZEI40` |
| `ZSD_QTY` | `ZQTY` |
| `ZSD_GWT` | `ZDE_GRWT` |
| `ZSD_NWT` | `ZDE_NET_WT` |
| `ZSD_HSN` | `ZNA_SEGMENT` |
| `ZSD_PACKDT`, `ZSD_REF_DT` | `ZSD_DATE` |
| `ZSD_EXPREF`, `ZSD_CNTY_ORGN` | `ZSD_UID` |
| `ZSD_PARTY_REF` | `ZREF` |
| `ZSD_PREC`, `ZSD_REC_PC`, `ZSD_POL`, `ZSD_POD`, `ZSD_PLD`, `ZSD_VSL_FLT` | `ZCHAR30` |
| `ZSD_PL_DECL`, `ZSD_INV_DECL`, `VBRK_VBELN` | `Z_STRING` |
| `ZSD_ADV_AMT` | `ZDE_AMT1` |

Data elements kept with their own name — **create only if they do not exist** (check SE11):

| Data element | Type | Len | Description |
|---|---|---:|---|
| `ZSD_NP` (for `ZSD_NP1`…`ZSD_NP5`) | `SSTRING` | 1333 | Notify party (multi-line) |
| `ZSD_MARK` / `ZSD_NOPACK` | `SSTRING` | 255 | Marks / No. of packages |
| `ZSD_CONT` / `ZSD_HSN_DESC` / `ZSD_CMMT` | `SSTRING` | 1333 | Containers / Description of goods / Commercial terms |

> ⚠ Check in SE11: `ZNA_SEGMENT` must hold at least 9 characters (HSN printed as `3917 2390`);
> `ZSERNO` / `ZSQNO` must be numeric or NUMC (used in arithmetic); `ZQTY`, `ZDE_GRWT`,
> `ZDE_NET_WT` are QUAN and `ZDE_AMT1` CURR — the table needs their unit / currency reference
> fields (§2). `ZSD_TART` (total articles) and `VBRK_VBELN` (all invoices, STRING) were not in the
> program; they use `ZSQNO` and `Z_STRING` — change if other data elements fit better.

Typing of the 6 data elements above: "Predefined type", field labels = description.
