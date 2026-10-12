# 08 — Message class `ZSD_EPACK` (SE91)

**Short text:** Export packing list (WRICEF 102-B). Create only if it does not exist.
Message 001 is the **exact FS text**. `&1`…`&4` are placeholders. The same list is in
`08_Message_Class_ZSD_EPACK.csv`.

| No. | Message text |
|---|---|
| 001 | Invalid Weight: Gross Weight cannot be less than Net Weight. |
| 002 | Package from Sl.No. &1: gross weight &2 kg is less than net weight &3 kg |
| 008 | Invoice &1 does not exist |
| 009 | Invoice &1 is cancelled or is a cancellation document |
| 011 | Invoices &3 differ in &1 (&2) - one packing list needs one value |
| 012 | No authorization: &1 &2, activity &3 |
| 015 | File &1 could not be read: &2 |
| 016 | Column "&1" not found in the heading row of the file |
| 017 | Row &1: mandatory field &2 is empty |
| 018 | Row &1: "&2" in column &3 is not a valid number |
| 019 | Row &1: "&2" in column &3 is not a valid date |
| 020 | Row &1: &2 differs from the value in row &3 |
| 021 | Row &1: Sl.No. &2 occurs more than once |
| 022 | Row &1: text in column &2 is longer than &3 characters |
| 027 | Row &1: first line item has no &2 (start of a package group) |
| 028 | Row &1: enter gross and net weight together |
| 029 | Sl.No. &1: quantity &2 differs from STD &3 x Articles &4 |
| 030 | Sl.No. &1: part no. &2 is not billed in the packing list invoices |
| 031 | Notify party &1 is filled but notify party &2 is empty |
| 032 | Packing list &1 written to ZSD_EPACK_DATA (&2 lines, &3 invoices) |
| 036 | The file contains no line items |
| 039 | Error &1 in &2 |
| 040 | Adobe Form &1 does not exist or is not active |
| 041 | Row &1: Sl.No. &2 is not in ascending order |
| 042 | No &1 partner in KNVP for customer &2, sales area &3 |
| 044 | Sl.No. &1: no HSN code in upload or material master |
| 047 | Packing list &1 has no valid invoices |
| 048 | Sl.No. &1: no product description in upload or material master |
| 049 | &1 not found - field stays empty on the output |
| 050 | &1 rows without Sl.No. were ignored (e.g. total row) |
| 052 | Packing list &1 could not be written to ZSD_EPACK_DATA |
| 055 | Row &1: no packing list number in or above this row - row ignored |
| 056 | &1 repeated invoice number(s) counted once |
| 057 | Row &1 is an exact duplicate of row &2 and was ignored |
| 058 | Invoice &1 is also assigned to packing list &2 in the same file |
| 059 | No packing list number found in the file |
| 060 | Customer &1 has &2 &3 partners in KNVP - &4 is used |
| 061 | Sales order &1 has no reference document - no order no. / date |
| 063 | The result list could not be displayed |
| 064 | Output of packing list &1 cancelled by the user |
| 065 | Form &1 output for packing list &2 |
| 066 | &1 of &2 packing lists output, &3 saved, &4 with errors |
| 068 | Row &1: same Article No. &2 as the row above but different &3 - new package |

**Total:** 43 messages.
