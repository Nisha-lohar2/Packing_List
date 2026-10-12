# WRICEF 102-B "Export Consolidated Packing List" — DDIC & repository objects

Package: **`ZSD_EPACK`**. All objects are new. No SAP object is modified,
appended or enhanced anywhere in this development.

Assumption references (`Axx`) point to `docs/Development_Assumptions.md`.

---

## 0. One Z table, as in the FS

The FS names one Z table, **`ZSD_EPACK_DATA`** (FS §2.1, `Z table Format1.xlsx` ›
Technical detail). It is the only Z table of this development. The report
writes the data of every packing list it has output into it: this is the log
of what was printed.

Two technical points differ from the Technical-detail sheet, and both are unavoidable in SAP:

| FS Technical detail | Implemented | Why |
|---|---|---|
| Key `ZSD_PACKNO` + `VBRK_VBELN` | Key `MANDT` + `ZSD_PACKNO` + `ZSD_SN` | `VBRK_VBELN` holds **all invoice numbers of the packing list in one row** (requested), so it is a `STRING` field, and a `STRING` cannot be a key field. Sl.No. makes each line item unique. |
| `VBRK_VBELN` CHAR 255 | `STRING`, e.g. `1000000348, 1000000349, 1000000350` | No length limit on the number of invoices |
| QUAN / CURR fields | + `ZSD_MEINS`, `ZSD_GEWEI`, `ZSD_WAERS` | DDIC requires a unit / currency reference field for QUAN / CURR fields |

Header fields are repeated on every line-item row, as in the FS layout.

---

## 1. Data elements — new

Data element name = field name (FS "Tech name").

| Data element | Type | Len | Dec | Description |
|---|---|---:|---:|---|
| `ZSD_PACKNO` | CHAR, conversion `ALPHA` | 10 | | Packing list number |
| `ZSD_SN` / `ZSD_INV_SN` | NUMC | 6 | | Sl.No. / Sr.No. INV |
| `ZSD_ARTNO` | CHAR | 15 | | Article No. |
| `ZSD_STD` / `ZSD_ART` / `ZSD_TART` | NUMC | 6 | | STD / Articles / Total Articles |
| `ZSD_PARTNO` / `ZSD_MATDESC` | CHAR | 40 | | Part No. / Product description |
| `ZSD_QTY` / `ZSD_GWT` / `ZSD_NWT` | QUAN | 13 | 3 | Quantity / Gross / Net weight |
| `ZSD_HSN` | CHAR | 17 | | H.S. code |
| `ZSD_PACKDT` / `ZSD_REF_DT` | DATS | 8 | | PL date / Buyer's ref. date |
| `ZSD_EXPREF` | CHAR | 20 | | Exporter ref. (Technical detail: "Not Req – fetched from SAP"; not filled) |
| `ZSD_PARTY_REF` | CHAR | 200 | | Buyer's reference no. |
| `ZSD_CNTY_ORGN` | CHAR | 20 | | Country of origin |
| `ZSD_NP` (for `ZSD_NP1`…`ZSD_NP5`) | `SSTRING` | 1333 | | Notify party (multi-line). FS: CHAR 200; the sample texts are longer (A19). |
| `ZSD_PREC`, `ZSD_REC_PC`, `ZSD_POL`, `ZSD_POD`, `ZSD_PLD`, `ZSD_VSL_FLT` | CHAR | 30 | | Shipping fields |
| `ZSD_MARK` / `ZSD_NOPACK` | `SSTRING` | 255 | | Marks / No. of packages (FS 30, multi-line) |
| `ZSD_CONT` / `ZSD_HSN_DESC` / `ZSD_CMMT` | `SSTRING` | 1333 | | Containers / Description of goods / Commercial terms |
| `ZSD_PL_DECL` / `ZSD_INV_DECL` | `STRING` | | | Declarations (FS CHAR 1500 exceeds the table limit) |
| `ZSD_ADV_AMT` | CURR | 15 | 2 | Advance amount |
| `ZSD_INVOICES` | `STRING` | | | All invoice numbers of the packing list (used for `VBRK_VBELN`) |
| `ZSD_EPACK_BLOCK` | CHAR | 10 | | Form text block id (form interface) |
| `ZSD_EPACK_ROW_TYPE` | CHAR, values `H` / `I` / `S` | 1 | | Printed row type (form interface) |
| `ZSD_EPACK_STATUS` | CHAR, values `S` / `W` / `E` | 1 | | Status in the result list |

---

## 2. Transparent table `ZSD_EPACK_DATA` — the only Z table (log)

Delivery class `A`, data class `APPL1`, size category `0`, no buffering,
log data changes yes, table maintenance not allowed (written only by `ZSD_EPACK`).

| # | Field | Key | Data element | Logic sheet | Excel column (FS template) |
|---|---|:---:|---|---|---|
| 1 | `MANDT` | ✔ | `MANDT` | | |
| 2 | `ZSD_PACKNO` | ✔ | `ZSD_PACKNO` | B9 | Packing list no |
| 3 | `ZSD_SN` | ✔ | `ZSD_SN` | B47 | Sl.No. |
| 4 | `VBRK_VBELN` | | `ZSD_INVOICES` (STRING) | input | SAP Invoice Number — all invoices, comma-separated |
| 5 | `ZSD_INV_SN` | | `ZSD_INV_SN` | B48 | Sr.No. INV |
| 6 | `ZSD_ARTNO` | | `ZSD_ARTNO` | B49 | Article No. |
| 7 | `ZSD_STD` | | `ZSD_STD` | B50 | STD |
| 8 | `ZSD_ART` | | `ZSD_ART` | B51 | Articles (first line of a package only) |
| 9 | `ZSD_TART` | | `ZSD_TART` | B52 | Total Articles — calculated, same on every row |
| 10 | `ZSD_PARTNO` | | `ZSD_PARTNO` | B53 | Part No. |
| 11 | `ZSD_MATDESC` | | `ZSD_MATDESC` | B54 | Product Description (as uploaded) |
| 12 | `ZSD_QTY` | | `ZSD_QTY` (ref. `ZSD_MEINS`) | B55 | Quantity |
| 13 | `ZSD_GWT` | | `ZSD_GWT` (ref. `ZSD_GEWEI`) | B56 | Gross Wt. (first line of a package only) |
| 14 | `ZSD_NWT` | | `ZSD_NWT` (ref. `ZSD_GEWEI`) | B57 | Net Wt. (first line of a package only) |
| 15 | `ZSD_HSN` | | `ZSD_HSN` | B58 | H.S Code (as uploaded) |
| 16 | `ZSD_PACKDT` | | `ZSD_PACKDT` | B10 | Packing list date |
| 17 | `ZSD_EXPREF` | | `ZSD_EXPREF` | B11 | Exporter Ref (not filled) |
| 18 | `ZSD_PARTY_REF` | | `ZSD_PARTY_REF` | B14 | Buyer's Reference No. |
| 19 | `ZSD_REF_DT` | | `ZSD_REF_DT` | B14 | Buyer's Reference date |
| 20 | `ZSD_CNTY_ORGN` | | `ZSD_CNTY_ORGN` | B16 | Country of origin |
| 21–25 | `ZSD_NP1`…`ZSD_NP5` | | `ZSD_NP` | B31 | Notifier Party 1–5 |
| 26 | `ZSD_PREC` | | `ZSD_PREC` | B33 | Pre-Carriage |
| 27 | `ZSD_REC_PC` | | `ZSD_REC_PC` | B34 | Place of Receipt by Pre-carrier |
| 28 | `ZSD_POL` | | `ZSD_POL` | B36 | Port of Loading |
| 29 | `ZSD_POD` | | `ZSD_POD` | B37 | Port of Discharge |
| 30 | `ZSD_PLD` | | `ZSD_PLD` | B38 | Place of Delivery |
| 31 | `ZSD_MARK` | | `ZSD_MARK` | B39 | Marks & Nos. |
| 32 | `ZSD_VSL_FLT` | | `ZSD_VSL_FLT` | B35 | Vessel / Flight No. |
| 33 | `ZSD_CONT` | | `ZSD_CONT` | B40 | Container No |
| 34 | `ZSD_NOPACK` | | `ZSD_NOPACK` | B41 | No. of Packages |
| 35 | `ZSD_HSN_DESC` | | `ZSD_HSN_DESC` | B42 | Description of goods |
| 36 | `ZSD_PL_DECL` | | `ZSD_PL_DECL` | B61 | Declaration of Packing list |
| 37 | `ZSD_ADV_AMT` | | `ZSD_ADV_AMT` (ref. `ZSD_WAERS`) | — | Advance amount received |
| 38 | `ZSD_CMMT` | | `ZSD_CMMT` | — | Commercial terms of invoice |
| 39 | `ZSD_INV_DECL` | | `ZSD_INV_DECL` | — | Declaration of Invoice |
| 40 | `ZSD_MEINS` | | `MEINS` (value `ST`) | | unit reference for `ZSD_QTY` |
| 41 | `ZSD_GEWEI` | | `GEWEI` (value `KG`) | | unit reference for weights |
| 42 | `ZSD_WAERS` | | `WAERS` | | currency reference for `ZSD_ADV_AMT` |

**When rows are written:** at the end of a run (not in "Validate only"), for every
packing list whose form was produced. One row per line item. If a packing list is
uploaded again, its earlier rows are deleted and the new ones written.
Packing lists with errors are not written; they appear only in the result list.

**Example** (sample file `test_data/Sample_data_ABAP.xlsx`): 9 rows, `ZSD_PACKNO`
`0090000088`, `ZSD_SN` 1–9, `VBRK_VBELN` = `1000000348, 1000000349, 1000000350`
on every row, `ZSD_TART` = 120.

---

## 3. Structures and table types — Smart Form interface

> For the **Adobe Form** layout (sample 90000014) see `src/Adobe_Form_Interface_Structures.md`:
> `ZSD_S_EPACK_AF_HDR`, `ZSD_S_EPACK_AF_ITM`, `ZSD_TT_EPACK_AF_ITM`, interface `ZSD_EPACK_AF_IF`.

Both Smart Forms have the **same** form interface:

| Parameter | Type | Content |
|---|---|---|
| `IS_HEADER` | `ZSD_S_EPACK_PRT_HDR` | Single values of the header |
| `IT_TEXT` | `ZSD_TT_EPACK_PRT_TXT` | All multi-line blocks (addresses, notify parties, …) |
| `IT_ITEM` | `ZSD_TT_EPACK_PRT_ITM` | Printed item rows incl. headings and subtotals |

### 3.1 `ZSD_S_EPACK_PRT_HDR`

| Component | Type | Content |
|---|---|---|
| `PACKNO` | `ZSD_PACKNO` | Packing List No. |
| `PACKDT_TXT` | CHAR 10 | Date, DD-MM-YYYY |
| `IEC` | CHAR 40 | Exporter's Ref. "IEC:…" |
| `GSTIN` | CHAR 18 | GST No. |
| `CIN` | CHAR 40 | CIN No. |
| `CNTY_ORGN` | CHAR 20 | Country of origin |
| `CNTY_DEST` | CHAR 50 | Country of final destination (ship-to) |
| `SAME_PARTY` | `XFELD` | Sold-to = ship-to → one "CONSIGNEE" block (A15, Q09) |
| `PREC`, `REC_PC`, `VSL_FLT`, `POL`, `POD`, `PLD` | CHAR 30 | Shipping fields |
| `NP_COUNT` | `INT1` | Number of filled notify parties (0–5) |
| `TOT_ART_TXT` | CHAR 10 | Total articles / packages |
| `TOT_QTY_TXT` | CHAR 17 | Total quantity |
| `TOT_GWT_TXT`, `TOT_NWT_TXT` | CHAR 17 | Total gross / net weight, 3 decimals |
| `TOT_TEXT` | CHAR 132 | "(Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages)" |

### 3.2 `ZSD_S_EPACK_PRT_TXT` / `ZSD_TT_EPACK_PRT_TXT` (standard table)

| Component | Type |
|---|---|
| `BLOCK` | `ZSD_EPACK_BLOCK` — `EXPORTER`, `SOLDTO`, `SHIPTO`, `NP1`…`NP5`, `ORDERREF`, `BUYERREF`, `PAYMENT`, `MARKS`, `CONTAINER`, `PACKAGES`, `GOODS`, `DECL` |
| `LINE_NO` | NUMC 3 |
| `TEXT` | CHAR 255 |

### 3.3 `ZSD_S_EPACK_PRT_ITM` / `ZSD_TT_EPACK_PRT_ITM` (standard table)

| Component | Type | Content |
|---|---|---|
| `ROW_TYPE` | `ZSD_EPACK_ROW_TYPE` | `H` / `I` / `S` |
| `ROW_TEXT` | CHAR 132 | Heading text or "<=Loose Pipes" |
| `SN_TXT`, `INV_SN_TXT`, `STD_TXT` | CHAR 6 | Without leading zeros |
| `ARTNO` | `ZSD_ARTNO` | |
| `ART_TXT` | CHAR 10 | Only on the first line of an article group |
| `PARTNO` | `ZSD_PARTNO` | |
| `MATDESC` | `ZSD_MATDESC` | |
| `QTY_TXT` | CHAR 17 | |
| `GWT_TXT`, `NWT_TXT` | CHAR 17 | Only on the first line of a weight group |
| `HSN` | `ZSD_HSN` | |
| `ART_FIRST`, `ART_LAST`, `WT_FIRST`, `WT_LAST` | `XFELD` | For the borders of the "merged" cells |
| `ART_GRP`, `WT_GRP` | NUMC 4 | |

---


## 4. Lock object `EZSD_EPACK` — new (SE11)

Primary table `ZSD_EPACK_DATA`, mode `E`, lock parameters `MANDT`, `ZSD_PACKNO`.
Generates `ENQUEUE_EZSD_EPACK` / `DEQUEUE_EZSD_EPACK` (parameter `MODE_ZSD_EPACK_DATA`).
It prevents two users writing the same packing list at the same time.

---

## 5. Message class `ZSD_EPACK` — new (SE91)

| No. | Text | Raised by |
|---|---|---|
| `001` | `Invalid Weight: Gross Weight cannot be less than Net Weight.` **(FS 2.1, exact text)** | rules |
| `002` | `Package from Sl.No. &1: gross weight &2 kg is less than net weight &3 kg` | rules |
| `008` | `Invoice &1 does not exist` | data |
| `009` | `Invoice &1 is cancelled or is a cancellation document` | data |
| `010` | `Billing type &2 of invoice &1 is not allowed for export packing lists` | data |
| `011` | `Invoices &3 differ in &1 (&2) - one packing list needs one value` | data |
| `012` | `No authorization: &1 &2, activity &3` | data / store |
| `013` | `Notify parties exist for packing list &1 - use format with notify party` | data |
| `014` | `Packing list &1 is locked by user &2 - not written to the log table` | save |
| `015` | `File &1 could not be read: &2` | upload |
| `016` | `Column "&1" not found in the heading row of the file` | upload |
| `017` | `Row &1: mandatory field &2 is empty` | upload |
| `018` | `Row &1: "&2" in column &3 is not a valid number` | upload |
| `019` | `Row &1: "&2" in column &3 is not a valid date` | upload |
| `020` | `Row &1: &2 differs from the value in row &3` | upload |
| `021` | `Row &1: Sl.No. &2 occurs more than once` | rules |
| `022` | `Row &1: text in column &2 is longer than &3 characters` | upload |
| `027` | `Row &1: first line item has no &2 (start of a package group)` | rules |
| `028` | `Row &1: enter gross and net weight together` | rules |
| `029` | `Sl.No. &1: quantity &2 differs from STD &3 x Articles &4` | rules (warning) |
| `030` | `Sl.No. &1: part no. &2 is not billed in the packing list invoices` | data (warning) |
| `031` | `Notify party &1 is filled but notify party &2 is empty` | upload |
| `032` | `Packing list &1 written to ZSD_EPACK_DATA (&2 lines, &3 invoices)` | save |
| `036` | `The file contains no line items` | upload |
| `038` | `PDF saved: &1` | report |
| `039` | `Error &1 in &2` | output |
| `040` | `Smart Form &1 does not exist or is not active` | output |
| `041` | `Row &1: Sl.No. &2 is not in ascending order` | rules |
| `042` | `No &1 partner in KNVP for customer &2, sales area &3` | data |
| `044` | `Sl.No. &1: no HSN code in upload or material master` | data (warning) |
| `047` | `Packing list &1 has no valid invoices` | data |
| `048` | `Sl.No. &1: no product description in upload or material master` | data (warning) |
| `049` | `&1 not found - field stays empty on the output` | data (warning) |
| `050` | `&1 rows without Sl.No. were ignored (e.g. total row)` | upload (info) |
| `052` | `Packing list &1 could not be written to ZSD_EPACK_DATA` | save |
| `053` | `File &1 could not be written` | output |
| `054` | `Enter a folder for the PDF download` | report |
| `055` | `Row &1: no packing list number in or above this row - row ignored` | upload |
| `056` | `&1 repeated invoice number(s) counted once` | upload (info) |
| `057` | `Row &1 is an exact duplicate of row &2 and was ignored` | upload (warning) |
| `058` | `Invoice &1 is also assigned to packing list &2 in the same file` | upload |
| `059` | `No packing list number found in the file` | upload / report |
| `060` | `Customer &1 has &2 &3 partners in KNVP - &4 is used` | data (warning) |
| `061` | `Sales order &1 has no reference document - no order no. / date` | data (warning) |
| `063` | `The result list could not be displayed` | output |
| `064` | `Output of packing list &1 cancelled by the user` | report (warning) |
| `065` | `Form &1 output for packing list &2` | report |
| `066` | `&1 of &2 packing lists output, &3 saved, &4 with errors` | report |
| `067` | `Validation only: &1 packing lists checked, &2 with errors` | report |
| `068` | `Row &1: same Article No. &2 as the row above but different &3 - new package` | rules (warning) |

---

## 6. Repository objects — new

| Object | Type | Purpose |
|---|---|---|
| `ZSD_EPACK` | Package | Container |
| `ZSD_EPACK_DATA` | Transparent table | **The only Z table** — log of printed packing lists (FS) |
| `EZSD_EPACK` | Lock object | Lock per packing list |
| `ZSD_S_EPACK_PRT_HDR`, `ZSD_S_EPACK_PRT_TXT`, `ZSD_S_EPACK_PRT_ITM` + table types | Structures | Smart Form interface (no database tables) |
| `ZSD_EPACK` | Message class | Messages |
| `ZSD_EPACK` | Authorization object | Print / write per company code |
| `ZSD_EPACK_PRINT` | Executable program | Report frame, START-OF-SELECTION |
| `ZSD_EPACK_PRINT_TOP` / `_SEL` / `_F01` / `_F02` | Includes | Types and data / selection screen / main FORMs / helper FORMs |
| `ZSD_EPACK_WITH_NP`, `ZSD_EPACK_WITHOUT_NP` | Smart Forms | See `docs/ZSD_EPACK_SmartForm_Specification.md` |
| `ZSD_EPACK_STYLE` | Smart Style | Shared formats |
| `ZSD_EPACK_CROSS` | SE78 graphic | Diagonal line for unused notify-party sections |
| **`ZSD_EPACK`** | Transaction (SE93) | Report transaction for `ZSD_EPACK_PRINT` (FS) |

---

## 7. Authorization (FS 2.3 is empty — A24)

| Object | Fields | Where |
|---|---|---|
| `S_TCODE` | `ZSD_EPACK` | Transaction start |
| **`ZSD_EPACK`** (new, SU21, class `SD`) | `ACTVT` 04 print / 01 write log, `BUKRS` | FORMs `CHECK_AUTHORITY`, `SAVE_PACKING_LISTS` |
| `V_VBRK_VKO` | `VKORG`, `ACTVT` = 04 | Print — per sales org. of the invoices |

---

## 8. Customizing / parameters

| Object | Entry | Purpose |
|---|---|---|
| `TVARVC` select-option `ZSD_EPACK_FKART` | Export tax-invoice billing types | Only these billing types are accepted. Empty = all accepted (A11). |
| `T001Z` party `ZCIN` | CIN per company code | Logic B7 |
| Company address field `BUILDING` | IEC number | Logic B11 |

---

## 9. SAP objects consumed — read only (logic sheet rows)

Every table is read **once per run** for all packing lists (FOR ALL ENTRIES).

| Table | Fields | Logic sheet |
|---|---|---|
| `VBRK` | `VBELN FKART FKSTO SFAKN BUKRS VKORG VTWEG SPART KUNAG ZTERM INCO1 INCO2 BUPLA` | B3, B6, B11, B20, B26, B43, B44 |
| `VBRP` | `VBELN POSNR MATNR WERKS AUBEL` | B13, B42 |
| `KNVP` | `KUNNR VKORG VTWEG SPART PARVW PARZA KUNN2 DEFPA` — `PARVW` `AG` (= SP), `WE` (= SH) | B20, B26 |
| `KNA1` | `KUNNR ADRNR` | B20, B26 |
| `T001` | `BUKRS ADRNR` | B3, B11 |
| `ADRC` | `NAME1 STREET STR_SUPPL1 CITY1 CITY2 POST_CODE1 REGION COUNTRY BUILDING` | B3, B11, B20, B26 |
| `ADR6` | `SMTP_ADDR` | B21, B27 |
| `ADR2` | `TELNR_LONG` where `R3_USER` = 1 (telephone) / 3 (mobile) | B22, B23, B28, B29 |
| `T005U` | `BEZEI` (logic sheet says `T005S`) | B4 |
| `T005T` | `LANDX` (logic sheet says `T005`) | B5, B20, B26 |
| `J_1BBRANCH` | `GSTIN` for `BUKRS` + `BRANCH` = `VBRK-BUPLA` | B6 |
| `T001Z` | `PAVAL` for `PARTY` = `ZCIN` | B7 |
| `VBAK` | `VGBEL` for `VBELN` = `VBRP-AUBEL` | B13 |
| `VBKD` | `BSTKD BSTDK` for `VBELN` = `VGBEL`, `POSNR` 000000 | B13 |
| `TVZBT` | `VTEXT` (SPRAS E, first `ZTAGG`) | B43 |
| `MAKT` | `MAKTX` (logic sheet says `MARA`) | B54 |
| `MARC` | `STEUC` for billed material + billing plant | B42, B58 |
| `T604N` | `TEXT1` for `SPRAS` E, `LAND1` IN | B42 |
| `TVARVC`, `USR01` | billing-type filter, default printer | — |

> ⚠ **Verify before activation** in DEV: `J_1BBRANCH-GSTIN`, `T604N` key and
> text field, `TVZBT-ZTAGG`, `KNVP-DEFPA`, `ADR2-R3_USER` values. Each read is
> in FORMs `READ_SAP_DATA` / `READ_ADDRESSES` (include `ZSD_EPACK_PRINT_F01`).

---

## 10. Activation order and transports

| Order | Objects |
|---|---|
| 1 | Data elements |
| 2 | Table `ZSD_EPACK_DATA`, lock object `EZSD_EPACK` |
| 3 | Form-interface structures and table types |
| 4 | Message class, authorization object |
| 5 | Smart Style, SE78 graphic, both Smart Forms |
| 6 | Program `ZSD_EPACK_PRINT` with includes `_TOP`, `_SEL`, `_F01`, `_F02`; transaction `ZSD_EPACK` |

Workbench request: everything above. Customizing: `TVARVC` entry, roles.
