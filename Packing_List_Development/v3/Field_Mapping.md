# Field mapping — Excel / ALV / ZSD_EPACK_DATA / Smart Form (version 3)

One line per Excel column, in the order of sheet **"Z table"** (B – AU → Excel A – AT).
**Src:** `SAP` = filled by report 1, protected in the ALV · `Man` = entered by the user, editable in the ALV.
Logic references are sheet *Data From SAP* / *Data From Excel* of the revised `Logic_sheet1.xlsx`.

| Excel | Heading (row 1) | Table field | Type | Src | Report 1 logic | Smart Form |
|---|---|---|---|---|---|---|
| A | Packing list no KEY FIELD | `ZSD_PACKNO` (key) | CHAR 10 | Man | — | `IS_HEADER-PACKNO` |
| B | SAP Sales Order No KEY FIELD | `VBAK_VBELN` (key) | CHAR 255 | SAP | `VBAK-VBELN` of the item (external format; stored internal, A05) | `IS_HEADER-SALES_ORDERS`, `IT_ITEMS-VBELN` |
| C | Sl.No. | `ZSD_SN` (key, A03) | NUMC 6 | Man | — | `IT_ITEMS-SN` |
| D | Sr.No. INV | `ZSD_INV_SN` | NUMC 6 | Man | — | `IT_ITEMS-INV_SN` |
| E | Article No. | `ZSD_ARTNO` | CHAR 15 | Man | — | `IT_ITEMS-ARTNO` |
| F | STD | `ZSD_STD` | NUMC 6 | Man | — | `IT_ITEMS-STD` |
| G | Articles | `ZSD_ART` | NUMC 6 | Man | — | `IT_ITEMS-ART`, sum `TOT_ART` |
| H | Total Articles | `ZSD_TART` | NUMC 6 | Man | — | `IS_HEADER-TART` |
| I | Part No. | `ZSD_PARTNO` | CHAR 40 | SAP | `VBAP-MATNR` (all POSNR), `CONVERSION_EXIT_MATN1_OUTPUT` | `IT_ITEMS-PARTNO` |
| J | Product Description | `ZSD_MATDESC` | CHAR 40 | SAP | `MAKT-MAKTX` (`SPRAS` E) for `VBAP-MATNR` (sheet: MARA, A11) | `IT_ITEMS-MATDESC` |
| K | Quantity      Pcs | `ZSD_QTY` | QUAN 13,3 | SAP | `VBAP-KWMENG` (unit `VBAP-VRKME` → `ZSD_MEINS`) | `IT_ITEMS-QTY`, sum `TOT_QTY` |
| L | Gross Wt. KGS | `ZSD_GWT` | QUAN 13,3 | Man | — | `IT_ITEMS-GWT`, sum `TOT_GWT` |
| M | Net Wt. KGS | `ZSD_NWT` | QUAN 13,3 | Man | — | `IT_ITEMS-NWT`, sum `TOT_NWT` |
| N | H.S Code | `ZSD_HSN` | CHAR 17 | SAP | `MARC-STEUC` (`VBAP-MATNR`, `VBAP-WERKS`) | `IT_ITEMS-HSN` |
| O | Exporter | `ZSD_EXP` | CHAR 100 | SAP | `VBAK-VKORG` → `TVKO-BUKRS` → `T001-ADRNR` → `ADRC` `NAME1, STREET, STR_SUPPL1, CITY1, POST_CODE1`, region `T005U-BEZEI`, country `T005T-LANDX`; fixed positions (A14) | `EXP_NAME` … `EXP_COUNTRY` |
| P | Packing list no | (= `ZSD_PACKNO`, A04) | CHAR 10 | Man | copy of column A | — |
| Q | Packing list date | `ZSD_PACKDT` | DATS | Man | — | `PACKDT` |
| R | Exporter ReF | `ZSD_EXPREF` | CHAR 20 | SAP | IEC No.: company `ADRC-BUILDING` (same path as Exporter) | `EXPREF` |
| S | CIN No. | `ZSD_CIN` | CHAR 25 | SAP | `T001Z-PAVAL`, `BUKRS`, `PARTY` = `ZCIN` | `CIN` |
| T | GST.No | `ZSD_GST` | CHAR 15 | SAP | `T001W-J_1BBRANCH` of `VBAP-WERKS` + `BUKRS` → `J_1BBRANCH-GSTIN` (A10) | `GST` |
| U | Order No | `ZSD_ORDNO` | CHAR 20 | SAP | `VBAK-VGBEL` → `VBKD-BSTKD` (`POSNR` 000000) — quotation customer reference | `ORDNO` |
| V | Order date | `ZSD_ORD_DATE` | DATS | SAP | `VBAK-VGBEL` → `VBKD-BSTDK` | `ORD_DATE` |
| W | Buyer's Reference No | `ZSD_PARTY_REF` | CHAR 200 | Man | — | `PARTY_REF` |
| X | Buyer's Reference date | `ZSD_REF_DT` | DATS | Man | — | `REF_DT` |
| Y | Country of origin of goods | `ZSD_CNTY_ORGN` | CHAR 20 | Man | — | `CNTY_ORGN` |
| Z | Country of final destination | `ZSD_CNTY_FIND` | CHAR 20 | SAP | ship-to `ADRC-COUNTRY` → `T005T-LANDX` | `CNTY_FIND` |
| AA | Sold to party | `ZSD_SOLD_PART` | CHAR 150 | SAP | `VBAK-KUNNR` → `KNA1-ADRNR` → `ADRC` `NAME1, STREET, STR_SUPPL1, CITY1, CITY2, POST_CODE1`, `T005T-LANDX`, `ADR6-SMTP_ADDR`, `ADR2-TELNR_LONG` (`R3_USER` 1 telephone, 3 mobile); fixed positions (A14) | `SOLD_*` |
| AB | Ship to Party | `ZSD_SHIP_PART` | CHAR 150 | SAP | `VBAK-KUNNR` + sales area → `KNVP` `PARVW` SH (`WE`) → `KUNN2` → `KNA1-ADRNR` → as sold-to (A13) | `SHIP_*` |
| AC – AG | Notifier Party1 – 5 | `ZSD_NP1` – `ZSD_NP5` | CHAR 200 | Man | — | `NP1` – `NP5`, `NP_COUNT` |
| AH | Pre-Carriage | `ZSD_PREC` | CHAR 30 | Man | — | `PREC` |
| AI | Place of Receipt by Pre-carrier | `ZSD_REC_PC` | CHAR 30 | Man | — | `REC_PC` |
| AJ | Port of Loading | `ZSD_POL` | CHAR 30 | Man | — | `POL` |
| AK | Port of Discharge | `ZSD_POD` | CHAR 30 | Man | — | `POD` |
| AL | Place of Delivery | `ZSD_PLD` | CHAR 30 | Man | — | `PLD` |
| AM | Terms of Payment | `ZSD_PTERMS` | CHAR 30 | SAP | `VBKD-ZTERM` (sales order, `POSNR` 000000) → `TVZBT-VTEXT` (E, first line) | `PTERMS` |
| AN | Terms of shipment | `ZSD_INCOT` | CHAR 20 | SAP | `VBKD-INCO1` + space + `VBKD-INCO2` (Q13) | `INCOT` |
| AO | Marks & Nos. | `ZSD_MARK` | CHAR 30 | Man | — | `MARK` |
| AP | Vessel / Flight No. | `ZSD_VSL_FLT` | CHAR 30 | Man | — | `VSL_FLT` |
| AQ | Container No | `ZSD_CONT` | CHAR 100 | Man | — | `CONT` |
| AR | No. of Packages | `ZSD_NOPACK` | CHAR 30 | Man | — | `NOPACK` |
| AS | Description of goods | `ZSD_HSN_DESC` | CHAR 100 | Man | — (A18) | `HSN_DESC` |
| AT | Declaration of packing list | `ZSD_PL_DECL` | STRING ≤ 1500 | Man | — | `PL_DECL` |

Not in the Excel file (technical, report 2): `MANDT`; `ZSD_MEINS` (`VBAP-VRKME` of the order item with the
same material, else `ST`); `ZSD_GEWEI` = `KG`; `ERNAM`, `ERDAT`, `ERZET` = user, date, time of saving.

## Report 1 — one ALV row per sales order item

Header fields (Exporter … Terms of shipment) are read once per sales order and repeated on each of its items.
SAP data is read in bulk before the rows are built (17 `SELECT`s, `FOR ALL ENTRIES` on unique keys,
hashed / sorted lookups). Missing data → field empty + warning in the message popup.

## Report 2 — Excel → table

Column *n* of the file is mapped to the field of position *n* above; the heading in row 1 must match
(normalized: upper case, letters and digits only). Each cell is converted to the type of its field
(`FORM text_to_field`): date, NUMC (whole number, max. digits), QUAN (number ≥ 0, 3 decimals),
CHAR / STRING (max. length). Excel row number is kept with every row.
