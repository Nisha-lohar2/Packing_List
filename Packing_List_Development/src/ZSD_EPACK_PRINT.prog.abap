*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_PRINT
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list"
*& Astral Limited / Project UDAY / SD
*& Started with SE38 / SA38 (no transaction code for now)
*&
*& 1. Upload the packing list Excel file (FS template "Z table
*&    Format1.xlsx" or the layout of Sample_data_ABAP.xlsx - columns are
*&    found by heading). One file may contain several packing lists.
*& 2. For every packing list: read the header data of its invoices and
*&    its line items exactly as described in the FS logic sheet
*&    (Logic_sheet1.xlsx).
*& 3. Output ONE Adobe Form (ZSD_EPACK_AF) per packing list number; the
*&    notify-party section shows only when notify parties are maintained.
*&    An error in one packing list never stops the others.
*& 4. At the end, write the data of every packing list that was output
*&    to the log table ZSD_EPACK_DATA (FS "Technical detail" layout: one
*&    row per line item, all invoice numbers in one field VBRK_VBELN).
*&
*& Procedural program - no custom classes or interfaces:
*&   ZSD_EPACK_PRINT_TOP  types, constants, global data
*&   ZSD_EPACK_PRINT_SEL  selection screen and its events
*&   ZSD_EPACK_PRINT_F01  main routines: upload, SAP data, forms, log table
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
*&   B01  Excel file
*& Selection texts
*&   P_FILE   Excel file (.xlsx)
*&---------------------------------------------------------------------*
