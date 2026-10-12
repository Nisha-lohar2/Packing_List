# Sample_data_ABAP.xlsx — expected results in `ZSD_EPACK_PRINT`

Test file provided by the business on 12.10.2026: one sheet ("Sheet2"),
headings in row 1, one packing list in rows 2–10.

## 1. How the file is read

| Item | In the file | How the report reads it |
|---|---|---|
| Column order | A = Packing list Number, B = SAP Invoice Number, C–M = item columns, N–AK = header columns (the "Technical detail" order, **not** the FS "Format" order) | Columns are found by their **heading**, so this order and the FS template order both work. Heading "Paclking list Number" (typo) is still recognised. |
| Packing list no. | 90000088 on every row | One packing list, stored as `0090000088` |
| Invoices | B2: "1000000348, 1000000349, 1000000350" | 3 invoices for packing list 90000088 |
| Package groups | "1 to 50" / 50 / 657 / 623 **repeated** on rows 2–4 (likewise rows 5–7 and 8–10), not merged | Repeated values = one package: counted once (A06) |
| "Fetch from logic" in H.S Code (M2), Exporter Ref (O2) and row 12 | Placeholder | Treated as empty. HSN comes from the material master (`MARC-STEUC`). Exporter Ref comes from SAP anyway. Row 12 is ignored (no Sl.No.). |
| Dates N2, Q2 | Real Excel dates | Read as stored: 10.09.2026 and 10.05.2026 (see §3, item 1) |
| Notify parties S–W | 5 filled (NP3–NP5 have the same text as NP2) | Format **with notify party** |

## 2. Expected result (computed offline with the parser rules)

| Check | Expected |
|---|---|
| Packing lists / forms | **1** packing list (90000088), **1** Adobe Form `ZSD_EPACK_AF` output, notify-party section shown |
| Invoices | 3 (1000000348, 1000000349, 1000000350) — must exist in SAP, not cancelled, same company code / business place / sold-to / sales area / payment terms / Incoterms |
| Line items | 9 (Sl.No. 1–9) |
| Packages | 3: "1 to 50" (50 articles, 657 / 623 kg), "51 to 90" (40, 540 / 510), "91 to 120" (30, 330 / 290) |
| Printed rows | Articles and weights on Sl.No. 1, 4 and 7 only; Sl.No. 2–3, 5–6 and 8–9 blank in those columns |
| Totals line | Articles **120** · "(Total 120 Packages)" · Qty **884** · Gross **1527.000** · Net **1423.000** |
| FS weight rule | Passed (657 ≥ 623, 540 ≥ 510, 330 ≥ 290) |
| HSN / description | HSN from `MARC-STEUC` of each Part No. in the billing plant. Description from column I. Section heading / description of goods "PVC Pipes - 39172390" when the material HSN is 39172390. |
| Warnings | 7 × message 029: quantity ≠ STD × Articles (Sl.No. 1, 2, 3, 4, 6, 7, 9). Not blocking. |
| Log table | After output: `ZSD_EPACK_DATA` **9 rows** (ZSD_PACKNO 90000088, ZSD_SN 1–9), header fields repeated on each row, `VBRK_VBELN` = "1000000348, 1000000349, 1000000350" on every row, `ZSD_TART` = 120. Message 032. Result status **W** (because of the warnings) |

## 3. Points for the business to check in this file

| # | Observation | Effect | Question |
|---|---|---|---|
| 1 | **Dates are probably swapped.** Q2 is stored as **10 May 2026**, but P2 says "Whatsapp 5th October 2026". N2 is stored as **10 Sep 2026** (displayed "09-10-26"), which may have been meant as 9 Oct 2026. The cells were typed day-first but Excel stored them month-first. | The form prints the stored date: Buyer's ref. date 10-05-2026, PL date 10-09-2026 | Enter dates as text `DD.MM.YYYY`, or check the Excel date format before uploading |
| 2 | Quantity ≠ STD × Articles on 7 of 9 rows (e.g. Sl.No. 1: 106 pcs vs 1 × 50). | 7 warnings (029) | Is "Qty = STD × Articles" a real rule? If not, the warning can be removed (Q12) |
| 3 | "No. of Packages" AF2 = 360, but the article totals add up to 120. | Header says 360, totals line says 120 | Which is correct? Is "Articles" the number of packages? |
| 4 | Part numbers F-1A, F-8A, FG00000000001/2/20 | Must be SAP materials, otherwise no HSN (warning 044) | Confirm they exist in the test system and are billed on the 3 invoices |
| 5 | Vessel "321321321", country of origin "India, Vietnam" | Printed as uploaded | Test values only? |
