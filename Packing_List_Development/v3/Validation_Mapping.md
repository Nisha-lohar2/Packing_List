# Validation mapping — report 2 ZSD_EPACK_UPLOAD (version 3)

Every message carries **Excel row**, **packing list no.**, **column (letter – heading)**, type and text.
Errors (E) block saving; warnings (W) do not. Save behaviour after errors: see V22.

| # | Check | Source | Level | Message | FORM |
|---|---|---|---|---|---|
| V01 | File path entered | Standard | Selection screen | (obligatory field) | `P_FILE OBLIGATORY` |
| V02 | Extension `.xlsx` | Requirement (format) | E, selection screen | Only Excel files .xlsx can be uploaded … | `check_file` |
| V03 | File exists on the PC | Requirement (path) | E, selection screen | File … does not exist | `check_file` |
| V04 | File readable, valid .xlsx, has a worksheet | Requirement (format) | E, stop | File … could not be read / not a valid Excel file / no worksheet | `read_xlsx` |
| V05 | Data below the heading row | Requirement | E, stop | The file has no data … | `read_file` |
| V06 | Heading of every column = layout heading at that position | Requirement (map by header / position) | E, stop | Heading "x" found, "y" expected … | `read_file` |
| V07 | Columns after the layout are empty | Requirement | W | Column … not part of the layout – ignored | `read_file` |
| V08 | Cell fits the field: valid date; whole number within NUMC length; number ≥ 0 for QUAN; text not longer than the field (STRING ≤ 1500) | Z-table sheet (types, lengths) | E | "x" is not a valid date / not a whole number / has more than n digits / is negative / Text has n characters, maximum m | `text_to_field` |
| V09 | 2nd "Packing list no" column empty or equal to the key column | Z-table sheet (A04) | E | "x" differs from the packing list no. of the key column … | `read_file` |
| V10 | Sales order value: each entry alphanumeric ≤ 10 characters; all together ≤ 255 | Z-table sheet (A05) | E | "x" is not a sales order number / Too many sales orders … | `normalize_orders` |
| V11 | Mandatory: Packing list no, SAP Sales Order No (key fields), Sl.No. (line key) | Z-table sheet "KEY FIELD", A03 | E | … is mandatory | `check_keys` |
| V12 | No duplicate key (packing list + sales order + Sl.No.) in the file | Table key | E | … also in Excel row n | `check_keys` |
| V13 | Sales order exists (`VBAK`) | Sheet *Validation* (VBAP / VBELN) | E | Sales order x does not exist | `check_orders` |
| V14 | Part No. is a valid material number | Sheet *Validation* (MATNR) | E | Part No. is not a valid material number | `check_orders` |
| V15 | Part No. is an item material of the sales order (`VBAP`) | Sheet *Validation* (VBAP VBELN, POSNR → MATNR) (Q11) | E | Part No. x is not an item of sales order y | `check_orders` |
| V16 | **Uploaded quantity ≤ delivered quantity**: Σ `ZSD_QTY` of all rows with the same sales order(s) + material ≤ Σ `LIPS-LFIMG` with `VGBEL` = sales order and the same `MATNR` (A06). Quantity > 0 without Part No. → error. | Sheet *Validation* | E (every row of the group) | **Uploaded Quantity should not be greater than Delivered Quantity** (material x: uploaded a, delivered b) | `check_orders` |
| V17 | **Gross weight ≥ net weight** per row | FS additional requirement | E | **Invalid Weight: Gross Weight cannot be less than Net Weight.** | `check_weights` |
| V18 | Same check again before each form output — no output for that packing list | FS ("Smart Form will not be output") | E | same text | `print_forms` |
| V19 | Format "without notifier party" chosen but notifier parties filled | FS 1.1 | E | Notifier party maintained – select format "with notifier party" | `check_packing_lists` |
| V20 | Packing list already in `ZSD_EPACK_DATA`: error unless "Overwrite" ticked (then warning) | No duplicate saves on rerun (A08) | E / W | Packing list is already saved – tick "Overwrite" … / … will be replaced | `check_packing_lists` |
| V21 | Header value differs from the first row of the packing list | A07 | W | Value differs from the first row … | `check_packing_lists` |
| V22 | Save: **all-or-nothing** (any E → nothing saved, nothing printed) or **per packing list** (packing lists without E saved) | Requirement — to confirm (Q01) | E | Nothing saved and printed: correct the errors … | `save_data` |
| V23 | Lock per packing list (`ENQUEUE_EZSD_EPACK`) | Standard | E | Packing list is locked by another user – not saved | `save_data` |
| V24 | Database error → `ROLLBACK WORK` of the LUW | Standard | E | Database error – not saved (rolled back) | `write_db` |
| V25 | Smart Form exists / print dialog not cancelled / output OK | Standard | E / W | Smart Form … not found / Printing cancelled … / output failed | `print_forms` |

Report 1 (ALV): the grid checks dates and numbers itself; `DATA_CHANGED` adds V17 as soon as both
weights of a row are entered, and the download is refused while marked entries exist.
