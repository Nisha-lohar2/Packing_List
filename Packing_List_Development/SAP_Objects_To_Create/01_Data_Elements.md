# 01 — Data elements (SE11)

Create **only the ones that do not exist**. For each line: SE11 → Data type → check whether a
`Z*` data element with the same type / length / decimals already exists (SE11 F4 or SE16 on
`DD04L`: `DATATYPE`, `LENG`, `DECIMALS`). If yes, reuse it and use its name in files 02–04.

Typing: "Predefined type" (no domain), except `ZSD_PACKNO` (domain `VBELN`).
Field labels: short = medium = long = heading, as below (shorten where SE11 requires).

| # | Data element | Type | Len | Dec | Short description | Field label | Reuse check | Note |
|---|---|---|---:|---:|---|---|---|---|
| 1 | `ZSD_PACKNO` | Domain VBELN (CHAR 10, ALPHA) |  |  | Packing list number | Packing List No. | ☐ exists? | table key; uses standard domain VBELN for ALPHA conversion |
| 2 | `ZSD_SN` | NUMC | 6 |  | Serial number | Sl.No. | ☐ exists? |  |
| 3 | `ZSD_INV_SN` | NUMC | 6 |  | Invoice serial number | Sr.No. INV | ☐ exists? |  |
| 4 | `ZSD_ARTNO` | CHAR | 15 |  | Article number | Article No. | ☐ exists? |  |
| 5 | `ZSD_STD` | NUMC | 6 |  | Standard packing | STD | ☐ exists? |  |
| 6 | `ZSD_ART` | NUMC | 6 |  | Articles | Articles | ☐ exists? |  |
| 7 | `ZSD_TART` | NUMC | 6 |  | Total articles | Total Articles | ☐ exists? |  |
| 8 | `ZSD_PARTNO` | CHAR | 40 |  | Part number | Part No. | ☐ exists? |  |
| 9 | `ZSD_MATDESC` | CHAR | 40 |  | Product description | Product Description | ☐ exists? |  |
| 10 | `ZSD_QTY` | QUAN | 13 | 3 | Quantity | Quantity | ☐ exists? | reference field ZSD_MEINS in the table |
| 11 | `ZSD_GWT` | QUAN | 13 | 3 | Gross weight | Gross Wt. | ☐ exists? | reference field ZSD_GEWEI in the table |
| 12 | `ZSD_NWT` | QUAN | 13 | 3 | Net weight | Net Wt. | ☐ exists? | reference field ZSD_GEWEI in the table |
| 13 | `ZSD_HSN` | CHAR | 17 |  | H.S. code | H.S Code | ☐ exists? |  |
| 14 | `ZSD_PACKDT` | DATS | 8 |  | Packing list date | PL Date | ☐ exists? |  |
| 15 | `ZSD_REF_DT` | DATS | 8 |  | Buyer's reference date | Ref. Date | ☐ exists? |  |
| 16 | `ZSD_EXPREF` | CHAR | 20 |  | Exporter reference | Exporter Ref. | ☐ exists? |  |
| 17 | `ZSD_PARTY_REF` | CHAR | 200 |  | Buyer's reference number | Buyer's Ref. | ☐ exists? |  |
| 18 | `ZSD_CNTY_ORGN` | CHAR | 20 |  | Country of origin | Ctry Origin | ☐ exists? |  |
| 19 | `ZSD_NP` | SSTRING | 1333 |  | Notify party | Notify Party | ☐ exists? | used by ZSD_NP1 … ZSD_NP5 |
| 20 | `ZSD_PREC` | CHAR | 30 |  | Pre-carriage | Pre-Carriage | ☐ exists? |  |
| 21 | `ZSD_REC_PC` | CHAR | 30 |  | Place of receipt by pre-carrier | Place Receipt | ☐ exists? |  |
| 22 | `ZSD_POL` | CHAR | 30 |  | Port of loading | Port Loading | ☐ exists? |  |
| 23 | `ZSD_POD` | CHAR | 30 |  | Port of discharge | Port Disch. | ☐ exists? |  |
| 24 | `ZSD_PLD` | CHAR | 30 |  | Place of delivery | Place Deliv. | ☐ exists? |  |
| 25 | `ZSD_VSL_FLT` | CHAR | 30 |  | Vessel / flight number | Vessel/Flight | ☐ exists? |  |
| 26 | `ZSD_MARK` | SSTRING | 255 |  | Marks and numbers | Marks & Nos. | ☐ exists? |  |
| 27 | `ZSD_NOPACK` | SSTRING | 255 |  | Number of packages | No. of Pkgs | ☐ exists? |  |
| 28 | `ZSD_CONT` | SSTRING | 1333 |  | Container numbers | Container No. | ☐ exists? |  |
| 29 | `ZSD_HSN_DESC` | SSTRING | 1333 |  | Description of goods | Descr. Goods | ☐ exists? |  |
| 30 | `ZSD_CMMT` | SSTRING | 1333 |  | Commercial terms of invoice | Comm. Terms | ☐ exists? |  |
| 31 | `ZSD_PL_DECL` | STRING |  |  | Declaration of packing list | PL Declaration | ☐ exists? |  |
| 32 | `ZSD_INV_DECL` | STRING |  |  | Declaration of invoice | Inv. Declaration | ☐ exists? |  |
| 33 | `ZSD_ADV_AMT` | CURR | 15 | 2 | Advance amount received | Advance Amt | ☐ exists? | reference field ZSD_WAERS in the table |
| 34 | `ZSD_INVOICES` | STRING |  |  | Invoice numbers of the packing list | Invoices | ☐ exists? | field VBRK_VBELN |

**Total:** 34 data elements. Fields of the Adobe structures (files 03 / 04) use
predefined types and standard data elements, so they need **no** further data elements.
