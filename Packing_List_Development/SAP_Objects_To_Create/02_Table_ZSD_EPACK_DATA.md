# 02 — Transparent table `ZSD_EPACK_DATA` (SE11) — the only Z table (log)

**Short description:** Export packing list — log of printed packing lists

| Attribute | Value |
|---|---|
| Delivery class | `A` (application table) |
| Data browser / table view maint. | Display / maintenance allowed with restrictions (no maintenance — written only by `ZSD_EPACK_PRINT`) |
| Technical settings | Data class `APPL1`, size category `0`, buffering not allowed, log data changes ✔ |
| Enhancement category | Can be enhanced (character-type or numeric) |

One row per line item; header fields repeated on every row; `VBRK_VBELN` holds **all
invoice numbers of the packing list** in one STRING (e.g. `1000000348, 1000000349, 1000000350`).
A STRING cannot be a key field, therefore the key is `MANDT` + `ZSD_PACKNO` + `ZSD_SN`.

| Field | Key | Data element | Currency / quantity reference | Content |
|---|:---:|---|---|---|
| `MANDT` | ✔ | `MANDT` |  |  |
| `ZSD_PACKNO` | ✔ | `ZSD_VBELN` |  | Packing list no |
| `ZSD_SN` | ✔ | `ZSERNO` |  | Sl.No. |
| `VBRK_VBELN` |  | `Z_STRING` |  | SAP Invoice Number — all invoices, comma-separated |
| `ZSD_INV_SN` |  | `ZSERNO` |  | Sr.No. INV |
| `ZSD_ARTNO` |  | `ZCHAR15` |  | Article No. |
| `ZSD_STD` |  | `ZSQNO` |  | STD |
| `ZSD_ART` |  | `ZSQNO` |  | Articles (first line of a package only) |
| `ZSD_TART` |  | `ZSQNO` |  | Total Articles — calculated, same on every row |
| `ZSD_PARTNO` |  | `ZSD_DEBEZEI40` |  | Part No. |
| `ZSD_MATDESC` |  | `ZSD_DEBEZEI40` |  | Product Description (as uploaded) |
| `ZSD_QTY` |  | `ZQTY` | `ZSD_MEINS` | Quantity |
| `ZSD_GWT` |  | `ZDE_GRWT` | `ZSD_GEWEI` | Gross Wt. (first line of a package only) |
| `ZSD_NWT` |  | `ZDE_NET_WT` | `ZSD_GEWEI` | Net Wt. (first line of a package only) |
| `ZSD_HSN` |  | `ZNA_SEGMENT` |  | H.S Code (as uploaded) |
| `ZSD_PACKDT` |  | `ZSD_DATE` |  | Packing list date |
| `ZSD_EXPREF` |  | `ZSD_UID` |  | Exporter Ref (not filled) |
| `ZSD_PARTY_REF` |  | `ZREF` |  | Buyer's Reference No. |
| `ZSD_REF_DT` |  | `ZSD_DATE` |  | Buyer's Reference date |
| `ZSD_CNTY_ORGN` |  | `ZSD_UID` |  | Country of origin |
| `ZSD_NP1` | | `ZSD_NP` | | Notify party 1 |
| `ZSD_NP2` | | `ZSD_NP` | | Notify party 2 |
| `ZSD_NP3` | | `ZSD_NP` | | Notify party 3 |
| `ZSD_NP4` | | `ZSD_NP` | | Notify party 4 |
| `ZSD_NP5` | | `ZSD_NP` | | Notify party 5 |
| `ZSD_PREC` |  | `ZCHAR30` |  | Pre-Carriage |
| `ZSD_REC_PC` |  | `ZCHAR30` |  | Place of Receipt by Pre-carrier |
| `ZSD_POL` |  | `ZCHAR30` |  | Port of Loading |
| `ZSD_POD` |  | `ZCHAR30` |  | Port of Discharge |
| `ZSD_PLD` |  | `ZCHAR30` |  | Place of Delivery |
| `ZSD_MARK` |  | `ZSD_MARK` |  | Marks & Nos. |
| `ZSD_VSL_FLT` |  | `ZCHAR30` |  | Vessel / Flight No. |
| `ZSD_CONT` |  | `ZSD_CONT` |  | Container No |
| `ZSD_NOPACK` |  | `ZSD_NOPACK` |  | No. of Packages |
| `ZSD_HSN_DESC` |  | `ZSD_HSN_DESC` |  | Description of goods |
| `ZSD_PL_DECL` |  | `Z_STRING` |  | Declaration of Packing list |
| `ZSD_ADV_AMT` |  | `ZDE_AMT1` | `ZSD_WAERS` | Advance amount received |
| `ZSD_CMMT` |  | `ZSD_CMMT` |  | Commercial terms of invoice |
| `ZSD_INV_DECL` |  | `Z_STRING` |  | Declaration of Invoice |
| `ZSD_MEINS` |  | `MEINS` |  | unit reference for `ZSD_QTY` |
| `ZSD_GEWEI` |  | `GEWEI` |  | unit reference for weights |
| `ZSD_WAERS` |  | `WAERS` |  | currency reference for `ZSD_ADV_AMT` |

**Currency / quantity fields tab:** `ZSD_QTY` → `ZSD_EPACK_DATA-ZSD_MEINS`;
`ZSD_GWT`, `ZSD_NWT` → `ZSD_EPACK_DATA-ZSD_GEWEI`; `ZSD_ADV_AMT` → `ZSD_EPACK_DATA-ZSD_WAERS`.

Check after activation: SE14 shows the table created in the database; total field count 42.
