*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_PRINT
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list"
*& Astral Limited / Project UDAY / SD
*& Transaction code : ZSD_EPACK  (as named in the FS)
*&
*& 1. Upload the packing list Excel file (FS template "Z table
*&    Format1.xlsx" or the layout of Sample_data_ABAP.xlsx - columns are
*&    found by heading). One file may contain several packing lists.
*& 2. For every packing list: read the header data of its invoices and
*&    its line items exactly as described in the FS logic sheet
*&    (Logic_sheet1.xlsx).
*& 3. Output ONE Smart Form per packing list number (with / without
*&    notify party). An error in one packing list never stops the others.
*& 4. At the end, save the data of every packing list that was output to
*&    the Z tables and write one log row per packing list (ZSD_EPACK_LOG).
*&
*& Option "Saved packing lists" reprints packing lists saved earlier,
*& selected by tax invoice and/or packing list number.
*&
*& Procedural program - no custom classes or interfaces:
*&   ZSD_EPACK_PRINT_TOP  types, constants, global data
*&   ZSD_EPACK_PRINT_SEL  selection screen and its events
*&   ZSD_EPACK_PRINT_F01  main routines: upload, SAP data, forms, save, log
*&   ZSD_EPACK_PRINT_F02  helper routines: package groups, weight rule,
*&                        totals, formatting, parsing
*&---------------------------------------------------------------------*
REPORT zsd_epack_print MESSAGE-ID zsd_epack.

INCLUDE zsd_epack_print_top.                " types, constants, global data
INCLUDE zsd_epack_print_sel.                " selection screen and its events

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  PERFORM main.

*----------------------------------------------------------------------*
* FORM routines
*----------------------------------------------------------------------*
INCLUDE zsd_epack_print_f01.                " upload, SAP data, forms, save, log
INCLUDE zsd_epack_print_f02.                " helper routines

*&---------------------------------------------------------------------*
*& Text elements
*&   B01  Data source
*&   B02  Output format
*&   B03  Output
*& Selection texts
*&   P_XLS    Upload Excel file
*&   P_FILE   Excel file (.xlsx)
*&   P_SAVED  Saved packing lists (reprint)
*&   S_VBELN  Tax Invoice
*&   S_PACKNO Packing List No.
*&   P_FAUTO  Automatic (notify party if maintained)
*&   P_FWITH  With notify party
*&   P_FWOUT  Without notify party
*&   P_PREV   Print preview
*&   P_PRINT  Print
*&   P_PDF    Download PDF (one file per packing list)
*&   P_DEST   Output device
*&   P_COPY   Number of copies
*&   P_DIR    PDF folder
*&   P_CHECK  Validate only (no form, no update)
*&---------------------------------------------------------------------*
