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

**When rows are written:** at the end of each run, for every
packing list whose form was produced. One row per line item. If a packing list is
uploaded again, its earlier rows are deleted and the new ones written.
Packing lists with errors are not written; they appear only in the result list.

**Example** (sample file `test_data/Sample_data_ABAP.xlsx`): 9 rows, `ZSD_PACKNO`
`0090000088`, `ZSD_SN` 1–9, `VBRK_VBELN` = `1000000348, 1000000349, 1000000350`
on every row, `ZSD_TART` = 120.

---

## 3. Adobe Form `ZSD_EPACK_AF` — interface, structures and table type

One Adobe Form for both FS formats. The notify-party section is shown only when notify
parties are maintained (`IS_HEADER-FORMAT` = `N`, `NP_COUNT` > 0). Full field lists,
examples and sources: **`src/Adobe_Form_Interface_Structures.md`**.

| Object | Type | Content |
|---|---|---|
| `ZSD_S_EPACK_AF_HDR` | Structure (SE11) | Header: 50 fields — exporter, consignee / buyer, notify parties, references, shipping, terms, marks / containers / packages, goods, totals, declaration |
| `ZSD_S_EPACK_AF_ITM` | Structure (SE11) | One printed row: row type, Sl.No., Sr.No. INV, Article No., STD, Articles, Part No., Product Description, Quantity, Gross / Net Wt., H.S Code, group fields |
| `ZSD_TT_EPACK_AF_ITM` | Table type (SE11) | Standard table of `ZSD_S_EPACK_AF_ITM` |
| `ZSD_EPACK_AF_IF` | Form interface (SFP) | `IS_HEADER` TYPE `ZSD_S_EPACK_AF_HDR`, `IT_ITEM` TYPE `ZSD_TT_EPACK_AF_ITM` |
| `ZSD_EPACK_AF` | Adobe Form (SFP) | Layout of samples 90000014 / 90000081 |

---

## 4. Message class `ZSD_EPACK` — new (SE91)


| No. | Text | Raised by |
|---|---|---|
| `001` | `Invalid Weight: Gross Weight cannot be less than Net Weight.` **(FS 2.1, exact text)** | rules |
| `002` | `Package from Sl.No. &1: gross weight &2 kg is less than net weight &3 kg` | rules |
| `008` | `Invoice &1 does not exist` | data |
| `009` | `Invoice &1 is cancelled or is a cancellation document` | data |
| `011` | `Invoices &3 differ in &1 (&2) - one packing list needs one value` | data |
| `012` | `No authorization: &1 &2, activity &3` | data |
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
| `039` | `Error &1 in &2` | output |
| `040` | `Adobe Form &1 does not exist or is not active` | output |
| `041` | `Row &1: Sl.No. &2 is not in ascending order` | rules |
| `042` | `No &1 partner in KNVP for customer &2, sales area &3` | data |
| `044` | `Sl.No. &1: no HSN code in upload or material master` | data (warning) |
| `047` | `Packing list &1 has no valid invoices` | data |
| `048` | `Sl.No. &1: no product description in upload or material master` | data (warning) |
| `049` | `&1 not found - field stays empty on the output` | data (warning) |
| `050` | `&1 rows without Sl.No. were ignored (e.g. total row)` | upload (info) |
| `052` | `Packing list &1 could not be written to ZSD_EPACK_DATA` | save |
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
| `068` | `Row &1: same Article No. &2 as the row above but different &3 - new package` | rules (warning) |

---

## 5. Repository objects — new (complete list)

| # | Object | Type | Transaction |
|---|---|---|---|
| 1 | `ZSD_EPACK` | Package | SE80 |
| 2 | Data elements of §1 | Data elements | SE11 |
| 3 | `ZSD_EPACK_DATA` | Transparent table — **the only Z table** (log of printed packing lists, FS) | SE11 |
| 4 | `ZSD_S_EPACK_AF_HDR` | Structure — Adobe Form header | SE11 |
| 5 | `ZSD_S_EPACK_AF_ITM` | Structure — Adobe Form item row | SE11 |
| 6 | `ZSD_TT_EPACK_AF_ITM` | Table type — Adobe Form items | SE11 |
| 7 | `ZSD_EPACK_AF_IF` | Adobe Form interface | SFP |
| 8 | `ZSD_EPACK_AF` | Adobe Form — the only form | SFP |
| 9 | `ZSD_EPACK` | Message class | SE91 |
| 10 | `ZSD_EPACK_PRINT` | Executable program (report frame) | SE38 |
| 11 | `ZSD_EPACK_PRINT_TOP` / `_SEL` / `_F01` / `_F02` | Includes: types and data / selection screen / main FORMs / helper FORMs | SE38 |

**Not created for now** (removed on request): transaction code, lock object,
custom authorization object, TVARVC billing-type filter, Smart Forms with their
structures, style and graphic. The report is started with SE38 / SA38.

---

## 6. Authorization

| Object | Fields | Where |
|---|---|---|
| `V_VBRK_VKO` (standard) | `VKORG`, `ACTVT` = 04 | FORM `CHECK_AUTHORITY` — print only for invoices of permitted sales organizations |

No custom authorization object for now (A24).

---

## 7. Master data used (no new object)

| Data | Entry | Purpose |
|---|---|---|
| `T001Z` party `ZCIN` | CIN per company code | Logic B7 |
| Company address field `BUILDING` | IEC number | Logic B11 |

---

## 8. SAP objects consumed — read only (logic sheet rows)


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

> ⚠ **Verify before activation** in DEV: `J_1BBRANCH-GSTIN`, `T604N` key and
> text field, `TVZBT-ZTAGG`, `KNVP-DEFPA`, `ADR2-R3_USER` values. Each read is
> in FORMs `READ_SAP_DATA` / `READ_ADDRESSES` (include `ZSD_EPACK_PRINT_F01`).

---

## 9. Activation order and transport

| Order | Objects |
|---|---|
| 1 | Package, data elements |
| 2 | Table `ZSD_EPACK_DATA` |
| 3 | Structures `ZSD_S_EPACK_AF_HDR`, `ZSD_S_EPACK_AF_ITM`, table type `ZSD_TT_EPACK_AF_ITM` |
| 4 | Form interface `ZSD_EPACK_AF_IF`, then Adobe Form `ZSD_EPACK_AF` |
| 5 | Message class `ZSD_EPACK` |
| 6 | Program `ZSD_EPACK_PRINT` with includes `_TOP`, `_SEL`, `_F01`, `_F02` |

One workbench request for everything above.
