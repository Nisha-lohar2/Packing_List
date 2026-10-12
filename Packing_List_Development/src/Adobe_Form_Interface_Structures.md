# WRICEF 102-B — Adobe Form `ZSD_EPACK_AF`: interface, structures and table type

Layout reference: `PackingList_90000014_Tarmac_UAE_BySea.pdf` (4 pages, format
**without notify party**, consignee block only). Every field printed on that form is
covered below. There is **one Adobe Form** for both FS formats: the notify-party section
(sample 90000081) is shown only when notify parties are maintained.

All objects are new, in package `ZSD_EPACK`. The report `ZSD_EPACK_PRINT` fills the
structures (FORM `BUILD_ONE`, `BUILD_PRINT_ITEMS`) and calls the form once per packing list
(FORM `SEND_FORM`: `FP_FUNCTION_MODULE_NAME`, `FP_JOB_OPEN`, generated function module,
`FP_JOB_CLOSE`). Logic-sheet rows (`Bxx`)
and Excel columns refer to `docs/Packing_List_Report_Design.md` §2.

| Object | Kind | Purpose |
|---|---|---|
| `ZSD_S_EPACK_AF_HDR` | Structure | Header: one value per field (everything above and below the item table) |
| `ZSD_S_EPACK_AF_ITM` | Structure | One printed row of the item table |
| `ZSD_TT_EPACK_AF_ITM` | Table type | Standard table of `ZSD_S_EPACK_AF_ITM` |
| `ZSD_EPACK_AF_IF` | Form interface (SFP, ABAP-Dictionary based) | Import `IS_HEADER` / `IT_ITEM` |
| `ZSD_EPACK_AF` | Adobe Form (SFP) | The only form — layout of samples 90000014 / 90000081 (§5) |

**Design rules**

* **Values are passed ready to print.** Numbers, weights, dates and the HSN are character fields
  already formatted by the report (as on the sample: `3398.310`, `27-05-2025`, `3917 2390`).
  The form needs no patterns or scripting. A value that must be blank on a row (Articles and weights
  on the 2nd+ line of a package) is simply empty.
* **Multi-line blocks are `STRING` fields** with line breaks (`cl_abap_char_utilities=>newline`).
  Bind them to multi-line text fields with "Allow multiple lines" and auto-fit height.
* **Section headings and subtotal rows are rows of `IT_ITEM`**, with `ROW_TYPE` choosing the row layout
  (Adobe: one table body row per row type, shown with a `presence` script on `ROW_TYPE`).
* **Totals are in the header**, because the totals row prints once at the end of the table.

---

## 1. Form interface `ZSD_EPACK_AF_IF`

| Parameter | Type | Pass by value | Optional |
|---|---|:---:|:---:|
| `IS_HEADER` | `ZSD_S_EPACK_AF_HDR` | ✔ | |
| `IT_ITEM` | `ZSD_TT_EPACK_AF_ITM` | ✔ | |

No global data and no initialization code are needed.

---

## 2. Header structure `ZSD_S_EPACK_AF_HDR`

Types: DE = existing data element; otherwise a predefined type (SE11 "Built-in type").
Example values are from packing list 90000014.

### 2.1 Document

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 1 | `PACKNO` | DE `ZSD_PACKNO` | 10 | Packing List No. | `90000014` | B9 · Excel M |
| 2 | `PACKDT` | CHAR | 10 | Date | `27-05-2025` | B10 · Excel N (DD-MM-YYYY) |
| 3 | `FORMAT` | CHAR | 1 | — (controls layout) | `O` | `N` = notify parties maintained → notify-party section shown; `O` = hidden (FS 1.1, A05) |
| 4 | `NP_COUNT` | INT1 | 3 | — | `0` | number of filled notify parties (0–5) |

### 2.2 Exporter (left box, "EXPORTER :")

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 5 | `EXP_NAME` | CHAR | 80 | EXPORTER : (1st line) | `ASTRAL LIMITED.` | B3 `ADRC-NAME1` of `T001-ADRNR` |
| 6 | `EXP_ADDRESS` | STRING | | address lines | `207/1, ASTRAL HOUSE, B/H RAJPATH CLUB,` ⏎ `OFF. S. G. HIGHWAY,` ⏎ `AHMEDABAD - 380 059, GUJARAT, INDIA.` | B3–B5 `ADRC` street / city / postcode, `T005U-BEZEI`, `T005T-LANDX` |
| 7 | `EXP_GSTIN` | CHAR | 18 | GST NO. (exporter box) and GST No. (right box) | `24AABCA2951N1ZO` | B6 `J_1BBRANCH-GSTIN` |
| 8 | `EXP_IEC` | CHAR | 20 | Exporter's Ref. → `IEC:` + value | `0897008171` | B11 `ADRC-BUILDING` |
| 9 | `EXP_CIN` | CHAR | 21 | CIN No. | `L25200GJ1996PLC029134` | B7 `T001Z-PAVAL` (`PARTY` = `ZCIN`) |
| 10 | `SIGN_FOR` | CHAR | 80 | Signature & Date → `For` + value | `Astral Limited` | company name `ADRC-NAME1` |

### 2.3 Consignee (ship-to) and buyer (sold-to)

The sample prints one "CONSIGNEE :" block because sold-to = ship-to (A15). When they differ,
the form prints both blocks (`SAME_PARTY` = space).

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 11 | `SAME_PARTY` | CHAR (`XFELD`) | 1 | — | `X` | sold-to = ship-to |
| 12 | `CNS_NAME` | CHAR | 80 | CONSIGNEE : (1st line) | `TARMAC BUILDING MATERIALS TRADING L.L.C` | B26 ship-to `ADRC-NAME1` |
| 13 | `CNS_ADDRESS` | STRING | | address lines | `P.O. BOX 36966, DUBAI, U.A.E.` | B26–B27 ship-to `ADRC`, `T005T-LANDX` |
| 14 | `CNS_CONTACT` | CHAR | 40 | `K/A:-` | `MR. TAREK. H` | **not in the Logic sheet** — empty until Q09 is answered |
| 15 | `CNS_TEL` | CHAR | 30 | `TEL.:-` | `+971 4 321 6364` | B28 ship-to `ADR2-TELNR_LONG` (`R3_USER` 1) |
| 16 | `CNS_FAX` | CHAR | 30 | `FAX :` | `+971 4 321 6374` | **not in the Logic sheet** — empty until Q09 is answered |
| 17 | `CNS_MOBILE` | CHAR | 30 | `MO :` | `+971 56 745 2764` | B28 ship-to `ADR2-TELNR_LONG` (`R3_USER` 3) |
| 18 | `CNS_EMAIL` | CHAR | 241 | `E-MAIL:` | `TAREK@TARMAC.ME` | B29 ship-to `ADR6-SMTP_ADDR` |
| 19 | `BUY_NAME` | CHAR | 80 | BUYER : (1st line) | — | B20 sold-to `ADRC-NAME1` |
| 20 | `BUY_ADDRESS` | STRING | | address lines | — | B20–B21 sold-to `ADRC`, `T005T-LANDX` |
| 21 | `BUY_TEL` | CHAR | 30 | `TEL.:-` | — | B22 sold-to `ADR2` (`R3_USER` 1) |
| 22 | `BUY_MOBILE` | CHAR | 30 | `MO :` | — | B22 sold-to `ADR2` (`R3_USER` 3) |
| 23 | `BUY_EMAIL` | CHAR | 241 | `E-MAIL:` | — | B23 sold-to `ADR6-SMTP_ADDR` |

### 2.4 Notify parties (shown only when `FORMAT` = `N`)

| # | Component | Type | Label on the form | Source |
|---|---|---|---|---|
| 24–28 | `NP1` … `NP5` | STRING | NOTIFY PARTY 1 … 5 | B31 · Excel S–W (`ZSD_NP1`…`ZSD_NP5`) |

### 2.5 References and countries (right box)

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 29 | `ORDER_REF` | STRING | | Order No. & Date | `PROFORMA INVOICE NO.: AL/TBMT/2025-26/MAY/01R1 & 03R4-04R4 DATE: 22.05 & 21.05` | B13 `VBRP-AUBEL` → `VBAK-VGBEL` → `VBKD-BSTKD`, `BSTDK` |
| 30 | `BUYER_REF` | STRING | | Buyer's Reference No. & Date | `E-MAIL DATED: SOBHA PROJECT PHASE 2 (3RD & 4TH SHIPMENT)` | B14 · Excel P + Q |
| 31 | `CNTY_ORGN` | CHAR | 20 | Country of Origin of Goods | `INDIA` | B16 · Excel R |
| 32 | `CNTY_DEST` | CHAR | 50 | Country of final Destination | `U.A.E` | B17 ship-to `T005T-LANDX` |

### 2.6 Shipping (left box, under the consignee)

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 33 | `PRE_CARR` | CHAR | 30 | Pre-Carriage | `BY ROAD/RAIL` | B33 · Excel X |
| 34 | `REC_PLACE` | CHAR | 30 | Place of Receipt by Pre-carrier | `ICD KHODIYAR, AHMEDABAD` | B34 · Excel Y |
| 35 | `VESSEL` | CHAR | 30 | Vessel / Flight No. | (empty) | B35 · Excel AD |
| 36 | `POL` | CHAR | 30 | Port of Loading | `MUNDRA SEAPORT, INDIA` | B36 · Excel Z |
| 37 | `POD` | CHAR | 30 | Port of Discharge | `JEBEL ALI SEAPORT, U.A.E` | B37 · Excel AA |
| 38 | `PLD` | CHAR | 30 | Place of Delivery | `JEBEL ALI SEAPORT, U.A.E` | B38 · Excel AB |

### 2.7 Terms (right box)

| # | Component | Type | Len | Label on the form | Example | Source |
|---|---|---|---:|---|---|---|
| 39 | `PAY_TERMS` | STRING | | Terms of Payment | `PAYMENT TERM: DOCUMENT AGAINST ACCEPTANCE (DA) -` ⏎ `90 DAYS AFTER B/L DATE` | B43 `VBRK-ZTERM` → `TVZBT-VTEXT` |
| 40 | `SHIP_TERMS` | CHAR | 60 | `TERMS OF SHIPMENT :-` + value | `EX-WORKS` | B44 `VBRK-INCO1` + `INCO2` (codes, Q10) |

### 2.8 Marks, containers, packages, goods

| # | Component | Type | Label on the form | Example | Source |
|---|---|---|---|---|---|
| 41 | `MARKS` | STRING | Marks & Nos. | (empty) | B39 · Excel AC |
| 42 | `CONTAINERS` | STRING | Container No. | `CSNU8464923-1X40' HC` ⏎ `TCLU1466959-1X40' HC` | B40 · Excel AE |
| 43 | `NO_PKGS` | STRING | No.of Pkgs. | `2342 Loose Pipes` ⏎ `1073 Boxes` ⏎ `3415 Pkgs.` | B41 · Excel AF; if empty, built from the package-type subtotals |
| 44 | `GOODS_DESC` | STRING | Description of Goods | `PP LOW NOISE PIPES (PP PIPES) - HS CODE : 3917 2390` ⏎ `PP LOW NOISE FITTINGS (PP FITTINGS) - HS CODE : 3917 4000` | B42 (Excel AG overrides, Q05) |

### 2.9 Totals row (end of the item table) and declaration

| # | Component | Type | Len | Printed in column | Example | Source |
|---|---|---|---:|---|---|---|
| 45 | `TOT_ART` | CHAR | 10 | Articles | `3415` | Σ Articles (B52) |
| 46 | `TOT_TEXT` | CHAR | 132 | Part No. + Product Description | `(Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages)` | built from the package-type subtotals |
| 47 | `TOT_QTY` | CHAR | 17 | Quantity Pcs | `39503` | Σ Quantity |
| 48 | `TOT_GWT` | CHAR | 17 | Gross Wt. KGS | `30409.155` | Σ Gross weight, 3 decimals |
| 49 | `TOT_NWT` | CHAR | 17 | Net Wt. KGS | `28934.410` | Σ Net weight, 3 decimals |
| 50 | `DECLARATION` | STRING | | Declaration : | `WE HEREBY CERTYFYING THAT` ⏎ `A. GROSS WEIGHT AND NET WEIGHT OF EACH ITEM IS SHOWN.` ⏎ … ⏎ `E. E-SEAL NO.: WIND02614414, WIND02614415` | B61 · Excel AH |

---

## 3. Item structure `ZSD_S_EPACK_AF_ITM`

One entry = one printed row of the table. The columns are those of the form, left to right.

| # | Component | Type | Len | Column on the form | Example (Sl.No. 1) | Source |
|---|---|---|---:|---|---|---|
| 1 | `ROW_TYPE` | CHAR | 1 | — (row layout) | `I` | `H` section heading · `I` line item · `S` package-type subtotal |
| 2 | `ROW_TEXT` | CHAR | 132 | heading text (`H`) / subtotal label (`S`) | (empty) | `H`: `PP LOW NOISE PIPES (PP PIPES) - HS CODE : 3917 2390` · `S`: `<=Loose Pipes` |
| 3 | `SL_NO` | CHAR | 6 | Sl.No. | `1` | B47 · Excel A |
| 4 | `INV_SN` | CHAR | 6 | Sr.No. INV | `3` | B48 · Excel B |
| 5 | `ARTICLE_NO` | CHAR | 15 | Article No. | `STACKER` | B49 · Excel C, as uploaded (a merged Excel cell gives the value on its first row only; Sl.No. 14: `1 TO 24`) |
| 6 | `STD` | CHAR | 6 | STD | `1` | B50 · Excel D |
| 7 | `ARTICLES` | CHAR | 10 | Articles | `171` | B51 · Excel E — first row of the article group only; `S` row: package-type subtotal (`2342`) |
| 8 | `PART_NO` | CHAR | 40 | Part No. | `M241270309` | B53 · Excel F |
| 9 | `PROD_DESC` | CHAR | 40 | Product Description | `110 MM LN GSS PIPE 3M` | B54 · Excel G, else `MAKT-MAKTX` |
| 10 | `QUANTITY` | CHAR | 17 | Quantity Pcs | `171` | B55 · Excel H |
| 11 | `GROSS_WT` | CHAR | 17 | Gross Wt. KGS | `3398.310` | B56 · Excel I — first row of the weight group only |
| 12 | `NET_WT` | CHAR | 17 | Net Wt. KGS | `3353.310` | B57 · Excel J — first row of the weight group only |
| 13 | `HSN` | CHAR | 10 | H.S Code | `3917 2390` | B58 · Excel K, else `MARC-STEUC`; printed "NNNN NNNN" |
| 14 | `PKG_TYPE` | CHAR | 20 | — | `Loose Pipes` | Excel AL (A09) |
| 15 | `ART_GRP` | NUMC | 4 | — | `0001` | article group number (rows sharing one Articles value = one package) |
| 16 | `ART_FIRST` | CHAR (`XFELD`) | 1 | — | `X` | first row of the article group |
| 17 | `ART_LAST` | CHAR (`XFELD`) | 1 | — | `X` | last row of the article group |
| 18 | `ART_ROWS` | INT2 | 5 | — | `1` | number of rows in the article group |
| 19 | `WT_GRP` | NUMC | 4 | — | `0001` | weight group number (rows sharing Gross / Net Wt.) |
| 20 | `WT_FIRST` | CHAR (`XFELD`) | 1 | — | `X` | first row of the weight group |
| 21 | `WT_LAST` | CHAR (`XFELD`) | 1 | — | ` ` | last row of the weight group (Sl.No. 3) |
| 22 | `WT_ROWS` | INT2 | 5 | — | `3` | number of rows in the weight group |

**"Merged" cells.** On the sample, Articles span several rows (e.g. Sl.No. 89–90, "1") and Gross / Net Wt. span three rows (Sl.No. 1–3, "3398.310 / 3353.310"). Adobe table rows
cannot be merged across rows. The value is printed on the first row of the group (`*_FIRST` = X,
other rows empty), and the horizontal border between rows of the group is hidden with a script on
`*_LAST` (bottom border only when `*_LAST` = X). `*_ROWS` is available if the cell should instead be
centred vertically over the group.

### Table type `ZSD_TT_EPACK_AF_ITM`

| Attribute | Value |
|---|---|
| Line type | `ZSD_S_EPACK_AF_ITM` |
| Access | Standard table |
| Key | Standard key, non-unique (not used — rows are printed in table order) |

### Row sequence for 90000014

| Rows of `IT_ITEM` | `ROW_TYPE` | Content |
|---|---|---|
| 1 | `H` | `PP LOW NOISE PIPES (PP PIPES) - HS CODE : 3917 2390` |
| 2–14 | `I` | Sl.No. 1–13 |
| 15 | `S` | `ARTICLES` = `2342`, `ROW_TEXT` = `<=Loose Pipes` |
| 16 | `H` | `PP LOW NOISE FITTINGS (PP FITTINGS) - HS CODE : 3917 4000` |
| 17–93 | `I` | Sl.No. 14–90 |
| 94 | `S` | `ARTICLES` = `1073`, `ROW_TEXT` = `<=Boxes` |

The totals row (3415 · text · 39503 · 30409.155 · 28934.410) comes from `IS_HEADER-TOT_*`, in the
table footer, last page only. The declaration and signature box print on the last page only.

---

## 4. Elements of the sample not taken from SAP or the upload

| Element on the sample | Handling |
|---|---|
| Title "PACKING LIST", all labels, column headings | Static text in the form |
| Page "Page 1 of 4" | Form system fields (current page / total pages) |
| Company stamp and signature | Not printed (signed by hand), or an MIME graphic if the business wants it |
| Diagonal line through the empty part of the table on the last page | Form layout (as on the sample; optional) |
| Header repeated on every page | Header subform placed on the master page |

---

## 5. Adobe Form `ZSD_EPACK_AF` (SFP) — layout notes

| Area | Content | Shown |
|---|---|---|
| Master page | Title "PACKING LIST", exporter box, consignee / buyer, notify parties, right box (PL no., date, exporter's ref., CIN, GST, order no., buyer's ref., countries, terms), shipping box, marks / containers / packages / goods, item column headings, page "Page n of m" | every page |
| Consignee | `SAME_PARTY` = X: one "CONSIGNEE :" block (`CNS_*`). Otherwise "SOLD TO / BILL TO PARTY (BUYER) :" (`BUY_*`) and "SHIP TO PARTY (CONSIGNEE) :" (`CNS_*`) | presence script on `SAME_PARTY` |
| Notify parties | Subform with "1ST … 5TH NOTIFY PARTY :" + `NP1`…`NP5`; each party subform hidden when its field is empty | presence script on `FORMAT` = `N` |
| Item table | Bound to `IT_ITEM`. Three body rows: heading (`ROW_TYPE` = `H`, `ROW_TEXT` across the table), item (`I`), subtotal (`S`, `ARTICLES` + `ROW_TEXT`) — each hidden unless its row type matches. Header row repeated on each page; rows may break between pages | every page |
| Totals row | `TOT_ART`, `TOT_TEXT`, `TOT_QTY`, `TOT_GWT`, `TOT_NWT` | table footer, last page |
| Declaration / signature | "Declaration :" + `DECLARATION`; "Signature & Date", "For " + `SIGN_FOR`, "Authorized Signatory" | last page |
