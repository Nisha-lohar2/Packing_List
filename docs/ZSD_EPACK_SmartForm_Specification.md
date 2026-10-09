# Smart Forms `ZSD_EPACK_WITH_NP` / `ZSD_EPACK_WITHOUT_NP` — build specification

WRICEF 102-B "Export consolidated packing list" — Astral Limited / UDAY / SD

Smart Forms are built in transaction `SMARTFORMS` and cannot be delivered
as source text. This page is the build instruction. The layout reproduces
the repository samples, which are treated as the approved format (A28):

* **without notify party** → `PackingList_90000014 Tarmac UAE_BySea.pdf`
* **with notify party** → `PackingList_90000081 Nabil Yemen_BySea - Rev.1.pdf`

The forms contain **no SELECT and no ABAP logic other than conditions**.
Every value is prepared by `ZCL_SD_EPACK_DATA` / `ZCL_SD_EPACK_RULES`.

---

## 1. Form interface (both forms, identical)

| Parameter | Type | Pass value |
|---|---|---|
| `IS_HEADER` | `ZSD_S_EPACK_PRT_HDR` | ✔ |
| `IT_TEXT` | `ZSD_TT_EPACK_PRT_TXT` | ✔ |
| `IT_ITEM` | `ZSD_TT_EPACK_PRT_ITM` | ✔ |

Global definitions: `GS_TEXT TYPE ZSD_S_EPACK_PRT_TXT`, `GS_ITEM TYPE ZSD_S_EPACK_PRT_ITM`.

Text blocks are printed with a loop node
`LOOP AT IT_TEXT INTO GS_TEXT WHERE BLOCK = '<block>'` and one text node `&GS_TEXT-TEXT&`.
Smart Forms wraps long lines at the window width.

---

## 2. Form attributes and style

| Attribute | Value |
|---|---|
| Page format | `DINA4`, portrait |
| Style | `ZSD_EPACK_STYLE` |
| Default paragraph | `P1` Arial 7 pt, single line spacing |
| Paragraphs | `P1` normal · `PB` bold · `PC` centred · `PR` right-aligned · `PH` header bold 12 pt centred ("PACKING LIST") · `PU` bold underlined (block captions "EXPORTER :", "1ST NOTIFY PARTY :") |
| Character formats | `B` bold · `U` underline · `S` small 6 pt (long notify-party address lines, as in sample 90000081) |
| Margins | left 1.0 cm, top 1.0 cm |

**One page `PAGE1`**, next page = `PAGE1`. All header windows are secondary
windows on `PAGE1`, so they repeat on every page (samples: identical
header on pages 1–4).

---

## 3. Windows — `ZSD_EPACK_WITHOUT_NP` (sample 90000014)

Coordinates in cm from the top-left page margin. The width of the usable area is 19.0 cm.
Exact positions are fine-tuned against the sample in the print preview.

| Window | Type | Left | Top | Width | Height | Content |
|---|---|---:|---:|---:|---:|---|
| `TITLE` | secondary | 0.0 | 0.0 | 19.0 | 0.6 | `PACKING LIST` (`PH`), frame |
| `EXPORTER` | secondary | 0.0 | 0.6 | 9.6 | 2.6 | "EXPORTER :" (`PU`) + block `EXPORTER` |
| `PLINFO` | secondary | 9.6 | 0.6 | 9.4 | 1.6 | Grid with 2 columns: "Packing List No. &IS_HEADER-PACKNO&" / "Date : &IS_HEADER-PACKDT_TXT&" / "CIN No." &IS_HEADER-CIN& — right: "Exporter's Ref." &IS_HEADER-IEC& / "GST No." &IS_HEADER-GSTIN& |
| `ORDERREF` | secondary | 9.6 | 2.2 | 9.4 | 0.6 | "Order No. & Date:" + block `ORDERREF` |
| `BUYERREF` | secondary | 9.6 | 2.8 | 9.4 | 0.8 | "Buyer's Reference No. & Date:" + block `BUYERREF` |
| `PARTIES` | secondary | 0.0 | 3.2 | 9.6 | 2.6 | see §5 |
| `COUNTRY` | secondary | 9.6 | 3.6 | 9.4 | 0.8 | 2 columns: "Country of Origin of Goods" / `&IS_HEADER-CNTY_ORGN&` — "Country of final Destination" / `&IS_HEADER-CNTY_DEST&` (`PC`) |
| `SHIPPING` | secondary | 0.0 | 5.8 | 9.6 | 1.6 | 2 × 3 grid: Pre-Carriage / Place of Receipt by Pre-carrier / Vessel / Flight No. / Port of Loading / Port of Discharge / Place of Delivery with `&IS_HEADER-PREC&` … `&IS_HEADER-PLD&` (`PC`) |
| `PAYMENT` | secondary | 9.6 | 4.4 | 9.4 | 3.0 | "Terms of Payment" + block `PAYMENT` |
| `MARKS` | secondary | 0.0 | 7.4 | 3.6 | 1.6 | "Marks & Nos." + block `MARKS`, then "Container No." (`U`) + block `CONTAINER` |
| `PACKAGES` | secondary | 3.6 | 7.4 | 6.0 | 1.6 | "No.of Pkgs." (`U`) + block `PACKAGES` |
| `GOODS` | secondary | 9.6 | 7.4 | 9.4 | 1.6 | "Description of Goods" (`U`) + block `GOODS` |
| `MAIN` | **main** | 0.0 | 9.0 | 19.0 | 16.4 | Item table, §6 |
| `FOOTER` | secondary | 0.0 | 25.4 | 19.0 | 2.4 | §7 — **condition: "Only after end of main window"** |
| `PAGENO` | secondary | 8.0 | 28.0 | 3.0 | 0.4 | `Page &SFSY-PAGE& of &SFSY-FORMPAGES&` (`PC`) |

All header windows have a 0.5 pt frame, as in the samples.

## 4. Windows — `ZSD_EPACK_WITH_NP` (sample 90000081)

Copy `ZSD_EPACK_WITHOUT_NP`, then change:

| Window | Left | Top | Width | Height | Content |
|---|---:|---:|---:|---:|---|
| `SOLDTO` | 0.0 | 3.2 | 9.6 | 1.8 | "SOLD TO / BILL TO PARTY (BUYER) :" (`PU`) + block `SOLDTO` |
| `SHIPTO` | 0.0 | 5.0 | 9.6 | 1.8 | "SHIP TO PARTY (CONSIGNEE) :" (`PU`) + block `SHIPTO` |
| `NP1` | 9.6 | 5.0 | 9.4 | 1.8 | "1ST NOTIFY PARTY :" + block `NP1` |
| `NP2` / `NP3` | 0.0 / 9.6 | 6.8 | 9.6 / 9.4 | 1.6 | "2ND …" / "3RD NOTIFY PARTY :" + blocks `NP2` / `NP3` |
| `NP4` / `NP5` | 0.0 / 9.6 | 8.4 | 9.6 / 9.4 | 1.6 | "4TH …" / "5TH NOTIFY PARTY :" + blocks `NP4` / `NP5` |
| `SHIPPING`, `PAYMENT` | | 10.0 | | 1.6 | moved down |
| `MARKS`, `PACKAGES`, `GOODS` | | 11.6 | | 2.0 | moved down; `CONTAINER` holds up to 5 lines |
| `MAIN` | 0.0 | 13.6 | 19.0 | 11.8 | about 21 item rows per page (sample: 21) |

**Unused notify parties (FS: "unused section will be crossed out")** — in each
`NPn` window:

* text node with condition `IS_HEADER-NP_COUNT >= n` → caption + block `NPn`
* graphic node `ZSD_EPACK_CROSS` (SE78, black and white, sized to the window, diagonal line), condition `IS_HEADER-NP_COUNT < n`

Smart Forms cannot draw lines dynamically. Because the NP windows have a fixed
size, a pre-drawn graphic of the same size gives the crossed-out look (A26).
Upload enforces consecutive notify parties (message 031), so the crossed-out
windows are always the last ones.

The `PARTIES` window of the "without" form is not used here.

## 5. `PARTIES` window (without-NP form) — sold-to / ship-to

* `IS_HEADER-SAME_PARTY = 'X'` → "CONSIGNEE :" + block `SHIPTO` (sample 90000014 shows one consignee block)
* otherwise → "SOLD TO / BILL TO PARTY (BUYER) :" + block `SOLDTO`, then "SHIP TO PARTY (CONSIGNEE) :" + block `SHIPTO` (FS #12, #13)

(A15 — confirm with CL-13.)

---

## 6. Main window — item table

Node `TABLE` over `IT_ITEM INTO GS_ITEM`. Total width 19.0 cm.

| # | Column heading | Width cm | Field | Align |
|---|---|---:|---|---|
| 1 | Sl.No. | 0.9 | `GS_ITEM-SN_TXT` | centre |
| 2 | Sr.No. INV | 0.9 | `GS_ITEM-INV_SN_TXT` | centre |
| 3 | Article No. | 2.1 | `GS_ITEM-ARTNO` | left |
| 4 | STD | 1.4 | `GS_ITEM-STD_TXT` | centre |
| 5 | Articles | 1.2 | `GS_ITEM-ART_TXT` | centre |
| 6 | Part No. | 2.4 | `GS_ITEM-PARTNO` | left |
| 7 | Product Description | 4.1 | `GS_ITEM-MATDESC` | left |
| 8 | Quantity Pcs | 1.4 | `GS_ITEM-QTY_TXT` | right |
| 9 | Gross Wt. KGS | 1.5 | `GS_ITEM-GWT_TXT` | right |
| 10 | Net Wt. KGS | 1.5 | `GS_ITEM-NWT_TXT` | right |
| 11 | H.S Code | 1.6 | `GS_ITEM-HSN` | centre |

### Line types

| Line type | Cells | Used for |
|---|---|---|
| `HEAD` | 11 | Column headings — **table header, "at page break" + "at start"** |
| `HSN` | 1 (19.0) | Section heading `&GS_ITEM-ROW_TEXT&`, bold underlined |
| `IT11` | 11 | Item, first line of article group **and** weight group |
| `IT10` | 11 | Item, first line of article group only |
| `IT01` | 11 | Item, first line of weight group only |
| `IT00` | 11 | Item, inside both groups |
| `SUBT` | 3 (5.3 / 1.2 / 12.5) | Package-type subtotal: empty / `&GS_ITEM-ART_TXT&` / `&GS_ITEM-ROW_TEXT&` |
| `TOTL` | 4 (5.3 / 1.2 / 7.9 / 4.6) | Table footer — see below |

Row selection (conditions on the table-row nodes):

| Condition | Line type |
|---|---|
| `GS_ITEM-ROW_TYPE = 'H'` | `HSN` |
| `GS_ITEM-ROW_TYPE = 'S'` | `SUBT` |
| `ROW_TYPE = 'I'` and `ART_FIRST = 'X'` and `WT_FIRST = 'X'` | `IT11` |
| `ROW_TYPE = 'I'` and `ART_FIRST = 'X'` and `WT_FIRST = ' '` | `IT10` |
| `ROW_TYPE = 'I'` and `ART_FIRST = ' '` and `WT_FIRST = 'X'` | `IT01` |
| `ROW_TYPE = 'I'` and `ART_FIRST = ' '` and `WT_FIRST = ' '` | `IT00` |

### "Merged" cells (sample: Article No. / Articles / Gross / Net span several lines)

Smart Forms tables cannot span rows. The effect is reproduced with borders:

* The data provider fills `ART_TXT`, `GWT_TXT` and `NWT_TXT` on the **first** line of a group only.
* In the table "Pattern" of each `ITxx` line type, columns 3 + 5 have a **top border only in `IT11`/`IT10`**, and columns 9 + 10 have a **top border only in `IT11`/`IT01`**. All other cells have full frames.
* The values print at the top of the group instead of vertically centred as in Excel (A27, CL-08).
* Groups may break across pages, as in sample 90000081 page 4 (A27).

### Table footer — line type `TOTL`, "at end of table" only

| Cell | Content |
|---|---|
| 1 | empty |
| 2 | `&IS_HEADER-TOT_ART_TXT&` (bold, underlined) |
| 3 | `&IS_HEADER-TOT_TEXT&` e.g. "(Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages)" |
| 4 | three sub-columns: `&IS_HEADER-TOT_QTY_TXT&` / `&IS_HEADER-TOT_GWT_TXT&` / `&IS_HEADER-TOT_NWT_TXT&` |

The samples cross out the empty rows below the last item with a diagonal line.
This cannot be done with a dynamic height. Instead, a text line
`*** END OF PACKING LIST ***` prints after the totals (A26, CL-38).

---

## 7. `FOOTER` window (last page only)

| Left part (13.0 cm) | Right part (6.0 cm, framed) |
|---|---|
| "Declaration :" (`U`) + block `DECL` (A–E lines incl. bottle-seal / e-seal numbers as uploaded) | "Signature & Date" / "For Astral Limited" / (space for stamp) / "Authorized Signatory" |

The company name in the signature box is the text "For Astral Limited". It can
also be printed from the first line of block `EXPORTER` if business wants it data-driven.

---

## 8. Test checklist for the forms

| # | Check | Data |
|---|---|---|
| 1 | Header repeats on every page; page n of m | 90 lines (sample 90000014) |
| 2 | 0 / 1 / 3 / 5 notify parties — unused windows crossed out | sample 90000081 with NP3–NP5 removed |
| 3 | Merged groups: values only on first line, inner borders open | sample 90000081 rows 1–4 |
| 4 | Three lines with own Articles but one weight | sample 90000014 lines 1–3 |
| 5 | HSN headings and "<=Loose Pipes" / "<=Boxes" subtotals | sample 90000014 |
| 6 | Totals 4074 / 12146 / 116323.800 / 115386.780 | full `Z table Format1.xlsx` |
| 7 | Declaration and signature only on the last page | any multi-page list |
| 8 | Same sold-to and ship-to → single "CONSIGNEE" block | sample 90000014 |
