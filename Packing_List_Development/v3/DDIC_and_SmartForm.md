# DDIC objects and Smart Form interface — WRICEF 102-B (version 3)

Source: sheet **"Z table"** of the revised `Logic_sheet1.xlsx` (copy in `v3/input/`).
Everything marked **(A..)** or **(Q..)** is explained in `Analysis_and_Open_Points.md`.

## 1. Objects

| # | Object | Type | New / reused |
|---|---|---|---|
| 1 | `ZDE_EPACK_*` (32) | Data elements | New (§2) |
| 2 | `ZSD_EPACK_DATA` | Transparent table | New (§3) — replaces the old invoice-based definition in `SAP_Objects_To_Create/02_…` |
| 3 | `EZSD_EPACK` | Lock object on `ZSD_EPACK_DATA` | New (§4) |
| 4 | `ZSD_S_EPACK_SF_HDR` | Structure — Smart Form header | New (§5) |
| 5 | `ZSD_S_EPACK_SF_ITM` | Structure — Smart Form item | New (§5) |
| 6 | `ZSD_TT_EPACK_SF_ITM` | Table type of `ZSD_S_EPACK_SF_ITM` | New (§5) |
| 7 | `ZSD_EPACK_SF` / `ZSD_EPACK_SF_NP` | Smart Forms without / with notifier party | New (§6) — names to confirm (Q09) |
| 8 | `ZSD_EPACK_COMMON` | Include program (type I) | New |
| 9 | `ZSD_EPACK_EXTRACT` | Report 1, T-code `ZSD_EPACK` | New |
| 10 | `ZSD_EPACK_UPLOAD` | Report 2, T-code `ZSD_EPACK_UPL` (to confirm, Q02) | New |
| 11 | GUI status `ZEPACK_STD` of `ZSD_EPACK_EXTRACT` | Copy of `SAPLKKBL` / `STANDARD_FULLSCREEN` + `ZDOWNLOAD` | New |

Reused SAP standard: `MANDT`, `MEINS`, `GEWEI`, `ERNAM`, `ERDAT`, `ERZET`, `INT4`; authorization object
`V_VBAK_VKO`; classes `CL_GUI_FRONTEND_SERVICES`, `CL_FDT_XL_SPREADSHEET`, `CL_SALV_TABLE`, `CL_BCS_CONVERT`;
FMs `REUSE_ALV_GRID_DISPLAY_LVC`, `LVC_FIELDCATALOG_MERGE`, `SSF_FUNCTION_MODULE_NAME`, `SSF_OPEN`, `SSF_CLOSE`.

If the existing S2A data elements from the earlier version (`ZSD_VBELN`, `ZSERNO`, `ZCHAR15`, `ZCHAR30` …)
have **exactly** the type and length below, they may be used instead of the new `ZDE_EPACK_*` ones.

## 2. Data elements (SE11, predefined type, field labels = description)

| Data element | Type | Len | Dec | Description (all labels) |
|---|---|---:|---:|---|
| `ZDE_EPACK_PACKNO` | CHAR | 10 | | Packing list no |
| `ZDE_EPACK_SO` | CHAR | 255 | | SAP Sales Order No |
| `ZDE_EPACK_SN` | NUMC | 6 | | Sl.No. |
| `ZDE_EPACK_INV_SN` | NUMC | 6 | | Sr.No. INV |
| `ZDE_EPACK_ARTNO` | CHAR | 15 | | Article No. |
| `ZDE_EPACK_STD` | NUMC | 6 | | STD |
| `ZDE_EPACK_ART` | NUMC | 6 | | Articles |
| `ZDE_EPACK_TART` | NUMC | 6 | | Total Articles |
| `ZDE_EPACK_PARTNO` | CHAR | 40 | | Part No. |
| `ZDE_EPACK_MATDESC` | CHAR | 40 | | Product Description |
| `ZDE_EPACK_QTY` | QUAN | 13 | 3 | Quantity Pcs |
| `ZDE_EPACK_GWT` | QUAN | 13 | 3 | Gross Wt. KGS |
| `ZDE_EPACK_NWT` | QUAN | 13 | 3 | Net Wt. KGS |
| `ZDE_EPACK_HSN` | CHAR | 17 | | H.S Code |
| `ZDE_EPACK_EXP` | CHAR | 100 | | Exporter |
| `ZDE_EPACK_PACKDT` | DATS | 8 | | Packing list date |
| `ZDE_EPACK_EXPREF` | CHAR | 20 | | Exporter Ref |
| `ZDE_EPACK_CIN` | CHAR | 25 | | CIN No. |
| `ZDE_EPACK_GST` | CHAR | 15 | | GST No. |
| `ZDE_EPACK_ORDNO` | CHAR | 20 | | Order No |
| `ZDE_EPACK_ORDDT` | DATS | 8 | | Order date |
| `ZDE_EPACK_PARTY_REF` | CHAR | 200 | | Buyer's Reference No |
| `ZDE_EPACK_REF_DT` | DATS | 8 | | Buyer's Reference date |
| `ZDE_EPACK_CNTY` | CHAR | 20 | | Country (origin / final destination) |
| `ZDE_EPACK_PARTY` | CHAR | 150 | | Party (sold-to / ship-to) |
| `ZDE_EPACK_NP` | CHAR | 200 | | Notifier Party |
| `ZDE_EPACK_C30` | CHAR | 30 | | Shipping text (30) |
| `ZDE_EPACK_INCOT` | CHAR | 20 | | Terms of shipment |
| `ZDE_EPACK_CONT` | CHAR | 100 | | Container No |
| `ZDE_EPACK_HSN_DESC` | CHAR | 100 | | Description of goods |
| `ZDE_EPACK_PL_DECL` | STRING | 1500 | | Declaration of packing list |
| `ZDE_EPACK_TEXT100` | CHAR | 100 | | Text (address line on the form) |

`ZDE_EPACK_PL_DECL`: the sheet says CHAR 1500; a table field of type CHAR is limited to 1333
characters, so STRING with maximum length 1500 is used (A15). The program also checks 1500.

## 3. Table `ZSD_EPACK_DATA`

Delivery class `A`, data browser "display/maintenance allowed with restrictions", data class `APPL1`,
size category `1`, no buffering, log data changes on. Enhancement category "can be enhanced (character-type or numeric)".

Order and names = sheet "Z table" (columns B – AU); sheet column Q (second "Packing list no",
`ZSD_PACKNO`) exists only in the Excel file (A04); `ZSD_VSl_FLT` is written `ZSD_VSL_FLT`.

| # | Field | Key | Data element | Ref. field | Sheet col. |
|---:|---|:---:|---|---|---|
| 1 | `MANDT` | ✔ | `MANDT` | | |
| 2 | `ZSD_PACKNO` | ✔ | `ZDE_EPACK_PACKNO` | | B (and Q) |
| 3 | `VBAK_VBELN` | ✔ | `ZDE_EPACK_SO` | | C |
| 4 | `ZSD_SN` | ✔ | `ZDE_EPACK_SN` | | D — **added to the key (A03)** |
| 5 | `ZSD_INV_SN` | | `ZDE_EPACK_INV_SN` | | E |
| 6 | `ZSD_ARTNO` | | `ZDE_EPACK_ARTNO` | | F |
| 7 | `ZSD_STD` | | `ZDE_EPACK_STD` | | G |
| 8 | `ZSD_ART` | | `ZDE_EPACK_ART` | | H |
| 9 | `ZSD_TART` | | `ZDE_EPACK_TART` | | I |
| 10 | `ZSD_PARTNO` | | `ZDE_EPACK_PARTNO` | | J |
| 11 | `ZSD_MATDESC` | | `ZDE_EPACK_MATDESC` | | K |
| 12 | `ZSD_QTY` | | `ZDE_EPACK_QTY` | `ZSD_EPACK_DATA-ZSD_MEINS` | L |
| 13 | `ZSD_GWT` | | `ZDE_EPACK_GWT` | `ZSD_EPACK_DATA-ZSD_GEWEI` | M |
| 14 | `ZSD_NWT` | | `ZDE_EPACK_NWT` | `ZSD_EPACK_DATA-ZSD_GEWEI` | N |
| 15 | `ZSD_HSN` | | `ZDE_EPACK_HSN` | | O |
| 16 | `ZSD_EXP` | | `ZDE_EPACK_EXP` | | P |
| 17 | `ZSD_PACKDT` | | `ZDE_EPACK_PACKDT` | | R |
| 18 | `ZSD_EXPREF` | | `ZDE_EPACK_EXPREF` | | S |
| 19 | `ZSD_CIN` | | `ZDE_EPACK_CIN` | | T |
| 20 | `ZSD_GST` | | `ZDE_EPACK_GST` | | U |
| 21 | `ZSD_ORDNO` | | `ZDE_EPACK_ORDNO` | | V |
| 22 | `ZSD_ORD_DATE` | | `ZDE_EPACK_ORDDT` | | W |
| 23 | `ZSD_PARTY_REF` | | `ZDE_EPACK_PARTY_REF` | | X |
| 24 | `ZSD_REF_DT` | | `ZDE_EPACK_REF_DT` | | Y |
| 25 | `ZSD_CNTY_ORGN` | | `ZDE_EPACK_CNTY` | | Z |
| 26 | `ZSD_CNTY_FIND` | | `ZDE_EPACK_CNTY` | | AA |
| 27 | `ZSD_SOLD_PART` | | `ZDE_EPACK_PARTY` | | AB |
| 28 | `ZSD_SHIP_PART` | | `ZDE_EPACK_PARTY` | | AC |
| 29–33 | `ZSD_NP1` … `ZSD_NP5` | | `ZDE_EPACK_NP` | | AD – AH |
| 34 | `ZSD_PREC` | | `ZDE_EPACK_C30` | | AI |
| 35 | `ZSD_REC_PC` | | `ZDE_EPACK_C30` | | AJ |
| 36 | `ZSD_POL` | | `ZDE_EPACK_C30` | | AK |
| 37 | `ZSD_POD` | | `ZDE_EPACK_C30` | | AL |
| 38 | `ZSD_PLD` | | `ZDE_EPACK_C30` | | AM |
| 39 | `ZSD_PTERMS` | | `ZDE_EPACK_C30` | | AN |
| 40 | `ZSD_INCOT` | | `ZDE_EPACK_INCOT` | | AO |
| 41 | `ZSD_MARK` | | `ZDE_EPACK_C30` | | AP |
| 42 | `ZSD_VSL_FLT` | | `ZDE_EPACK_C30` | | AQ |
| 43 | `ZSD_CONT` | | `ZDE_EPACK_CONT` | | AR |
| 44 | `ZSD_NOPACK` | | `ZDE_EPACK_C30` | | AS |
| 45 | `ZSD_HSN_DESC` | | `ZDE_EPACK_HSN_DESC` | | AT |
| 46 | `ZSD_PL_DECL` | | `ZDE_EPACK_PL_DECL` | | AU |
| 47 | `ZSD_MEINS` | | `MEINS` | | technical — unit of `ZSD_QTY` (A15) |
| 48 | `ZSD_GEWEI` | | `GEWEI` | | technical — unit of the weights, `KG` (A15) |
| 49 | `ERNAM` | | `ERNAM` | | technical — saved by |
| 50 | `ERDAT` | | `ERDAT` | | technical — saved on |
| 51 | `ERZET` | | `ERZET` | | technical — saved at |

Key length: 3 + 10 + 255 + 6 = 274 characters (limit 900 bytes — OK).
**SAP-VERIFY:** the character fields add up to about 2 700 characters. If activation reports that the
table is too wide, change `ZDE_EPACK_NP` (5 × 200) to `SSTRING` 200 — no program change needed.

## 4. Lock object `EZSD_EPACK`

Primary table `ZSD_EPACK_DATA`, lock mode `E`, lock parameters **`MANDT`, `ZSD_PACKNO`** (remove
`VBAK_VBELN` and `ZSD_SN` from the parameters). Generates `ENQUEUE_EZSD_EPACK` / `DEQUEUE_EZSD_EPACK`.

## 5. Smart Form interface structures

**`ZSD_S_EPACK_SF_HDR`** — one per packing list (filled in FORM `form_data` of report 2)

| Component | Type | Content |
|---|---|---|
| `PACKNO` | `ZDE_EPACK_PACKNO` | Packing list no |
| `PACKDT` | `ZDE_EPACK_PACKDT` | Packing list date |
| `SALES_ORDERS` | `ZDE_EPACK_SO` | Sales orders of the packing list, `, ` separated |
| `EXP_NAME`, `EXP_STREET`, `EXP_STREET2`, `EXP_CITY`, `EXP_POSTCODE`, `EXP_REGION`, `EXP_COUNTRY` | `ZDE_EPACK_TEXT100` | Exporter, positions 1–7 of `ZSD_EXP` |
| `EXPREF` | `ZDE_EPACK_EXPREF` | Exporter Ref (IEC) |
| `CIN` | `ZDE_EPACK_CIN` | |
| `GST` | `ZDE_EPACK_GST` | |
| `ORDNO` | `ZDE_EPACK_ORDNO` | Order no. |
| `ORD_DATE` | `ZDE_EPACK_ORDDT` | Order date |
| `PARTY_REF` | `ZDE_EPACK_PARTY_REF` | Buyer's reference no. |
| `REF_DT` | `ZDE_EPACK_REF_DT` | Buyer's reference date |
| `CNTY_ORGN` | `ZDE_EPACK_CNTY` | Country of origin |
| `CNTY_FIND` | `ZDE_EPACK_CNTY` | Country of final destination |
| `SOLD_NAME`, `SOLD_STREET`, `SOLD_STREET2`, `SOLD_CITY`, `SOLD_DISTRICT`, `SOLD_POSTCODE`, `SOLD_COUNTRY`, `SOLD_EMAIL`, `SOLD_TEL`, `SOLD_MOBILE` | `ZDE_EPACK_TEXT100` | Sold-to, positions 1–10 of `ZSD_SOLD_PART` |
| `SHIP_NAME` … `SHIP_MOBILE` (same 10) | `ZDE_EPACK_TEXT100` | Ship-to (consignee), positions 1–10 of `ZSD_SHIP_PART` |
| `NP1` … `NP5` | `ZDE_EPACK_NP` | Notifier parties |
| `NP_COUNT` | `INT4` | Number of filled notifier parties (unused ones crossed out) |
| `PREC`, `REC_PC`, `VSL_FLT`, `POL`, `POD`, `PLD`, `PTERMS`, `MARK`, `NOPACK` | `ZDE_EPACK_C30` | Shipment texts |
| `INCOT` | `ZDE_EPACK_INCOT` | Terms of shipment |
| `CONT` | `ZDE_EPACK_CONT` | Container no. |
| `HSN_DESC` | `ZDE_EPACK_HSN_DESC` | Description of goods |
| `PL_DECL` | `ZDE_EPACK_PL_DECL` | Declaration |
| `TART` | `ZDE_EPACK_TART` | Total Articles as entered |
| `TOT_ART` | `INT4` | Sum of Articles |
| `TOT_QTY` | `ZDE_EPACK_QTY` (ref. `MEINS`) | Sum of quantity |
| `TOT_GWT`, `TOT_NWT` | `ZDE_EPACK_GWT` / `ZDE_EPACK_NWT` (ref. `GEWEI`) | Sum of weights |
| `MEINS` | `MEINS` | Unit of quantity |
| `GEWEI` | `GEWEI` | Unit of weight (`KG`) |

**`ZSD_S_EPACK_SF_ITM`** — one per row, sorted by Sl.No.

| Component | Type | Component | Type |
|---|---|---|---|
| `SN` | `ZDE_EPACK_SN` | `MATDESC` | `ZDE_EPACK_MATDESC` |
| `INV_SN` | `ZDE_EPACK_INV_SN` | `QTY` | `ZDE_EPACK_QTY` (ref. `MEINS`) |
| `ARTNO` | `ZDE_EPACK_ARTNO` | `MEINS` | `MEINS` |
| `STD` | `ZDE_EPACK_STD` | `GWT`, `NWT` | `ZDE_EPACK_GWT` / `_NWT` (ref. `GEWEI`) |
| `ART` | `ZDE_EPACK_ART` | `GEWEI` | `GEWEI` |
| `PARTNO` | `ZDE_EPACK_PARTNO` | `HSN` | `ZDE_EPACK_HSN` |
| | | `VBELN` | `ZDE_EPACK_SO` |

**`ZSD_TT_EPACK_SF_ITM`**: standard table of `ZSD_S_EPACK_SF_ITM`, default key.

## 6. Smart Forms `ZSD_EPACK_SF` (without NP) and `ZSD_EPACK_SF_NP` (with NP)

Form interface (identical for both), import parameters:

| Parameter | Type | Pass value |
|---|---|---|
| `IS_HEADER` | `ZSD_S_EPACK_SF_HDR` | ✔ |
| `IT_ITEMS` | `ZSD_TT_EPACK_SF_ITM` | ✔ |

Called by report 2 with `SSF_FUNCTION_MODULE_NAME` → `SSF_OPEN` → generated FM per packing list
(`CONTROL_PARAMETERS-NO_OPEN` / `NO_CLOSE` = X, one spool request) → `SSF_CLOSE`.
The format is chosen by the user on the selection screen (FS 1.1); "without NP" is refused when
notifier parties are filled (validation V19).

Layout (from the sample PDFs `PackingList_90000014 …` without NP and `PackingList_90000081 …` with NP):

| Window | Fields |
|---|---|
| Logo / title "PACKING LIST" | graphic |
| Exporter | `EXP_*` lines (empty lines suppressed), `GST`, `CIN` |
| Packing list no. & date | `PACKNO`, `PACKDT` |
| Exporter Ref | `EXPREF` |
| Order no. & date | `ORDNO`, `ORD_DATE` |
| Buyer's reference no. & date | `PARTY_REF`, `REF_DT` |
| Consignee (ship-to) | `SHIP_*`; labels "E-mail", "Telephone", "Mobile number" printed only when filled |
| Buyer (sold-to) | `SOLD_*` as above |
| Notifier parties (form `_NP` only) | `NP1`–`NP5`; slots `> NP_COUNT` crossed out |
| Countries | `CNTY_ORGN`, `CNTY_FIND` |
| Shipment | `PREC`, `REC_PC`, `VSL_FLT`, `POL`, `POD`, `PLD`, `PTERMS`, `INCOT` |
| Marks / containers / packages / goods | `MARK`, `CONT`, `NOPACK`, `HSN_DESC` |
| Main window — table | loop `IT_ITEMS`: `SN`, `INV_SN`, `ARTNO`, `STD`, `ART`, `PARTNO`, `MATDESC`, `QTY`, `GWT`, `NWT`, `HSN` (empty Articles / weights = merged with the package above, as on the samples) |
| Totals line | `TOT_ART` (or `TART` when entered — Q06), `TOT_QTY`, `TOT_GWT`, `TOT_NWT` |
| Footer | `PL_DECL`, signature |

Exact window positions follow the approved sample PDFs (FS: "Development will be carried out based on
the approved Packing List format provided").
