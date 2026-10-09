*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_PRINT
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list"
*& Astral Limited / Project UDAY / SD
*& Transaction code : ZSD_EPACK  (as named in the FS)
*&
*& Prints the export packing list for one or more tax invoices with the
*& Smart Form "with notify party" or "without notify party".
*&
*& The program is only a shell: selection screen and delegation.
*&   ZCL_SD_EPACK_DATA    - reads Z tables + SAP data, all validations
*&   ZCL_SD_EPACK_RULES   - weight check, totals, formatting (unit-tested)
*&   ZCL_SD_EPACK_OUTPUT  - Smart Form call, PDF, message popup
*& Data is maintained with transaction ZSD_EPACK_UPL (ZSD_EPACK_UPLOAD).
*&---------------------------------------------------------------------*
REPORT zsd_epack_print.

DATA gv_vbeln TYPE vbeln_vf.

*&---------------------------------------------------------------------*
*& Selection screen
*&   FS / Logic_sheet1.xlsx "Input": Tax Invoice From - To, multiple
*&   selection. Everything else is ASSUMPTION A03, A05, A25.
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
SELECT-OPTIONS s_vbeln FOR gv_vbeln.
PARAMETERS     p_packno TYPE zsd_packno.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-b02.
PARAMETERS p_fauto RADIOBUTTON GROUP fmt DEFAULT 'X'.
PARAMETERS p_fwith RADIOBUTTON GROUP fmt.
PARAMETERS p_fwout RADIOBUTTON GROUP fmt.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-b03.
PARAMETERS p_prev  RADIOBUTTON GROUP out DEFAULT 'X' USER-COMMAND out.
PARAMETERS p_print RADIOBUTTON GROUP out.
PARAMETERS p_pdf   RADIOBUTTON GROUP out.
PARAMETERS p_dest  TYPE rspopname.
PARAMETERS p_copy  TYPE tdcopies DEFAULT 1.
PARAMETERS p_path  TYPE rlgrap-filename LOWER CASE MODIF ID pdf.
SELECTION-SCREEN END OF BLOCK b3.

*&---------------------------------------------------------------------*
*& Local controller
*&---------------------------------------------------------------------*
CLASS lcl_report DEFINITION FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS run.

    CLASS-METHODS pdf_file_dialog
      CHANGING cv_path TYPE rlgrap-filename.

  PRIVATE SECTION.
    METHODS format
      RETURNING VALUE(rv_format) TYPE char1.

    METHODS output
      RETURNING VALUE(rv_output) TYPE char1.

ENDCLASS.


CLASS lcl_report IMPLEMENTATION.

  METHOD run.

    DATA(lo_data) = NEW zcl_sd_epack_data( ).

    TRY.
        IF p_pdf = abap_true AND p_path IS INITIAL.
          zcx_sd_epack=>raise( iv_msgno = '054' ).
        ENDIF.

        DATA(lv_packno) = lo_data->resolve_packno( it_vbeln  = s_vbeln[]
                                                   iv_packno = p_packno ).

        DATA(ls_print) = lo_data->get_print_data( iv_packno = lv_packno
                                                  iv_format = format( ) ).

      CATCH zcx_sd_epack INTO DATA(lx_error).
        " FS 2.1: on an error "smart form will not come in output"
        IF lx_error->mt_messages IS NOT INITIAL.
          zcl_sd_epack_output=>show_messages( it_messages = lx_error->mt_messages
                                              iv_title    = 'Packing list cannot be printed' ).
        ENDIF.
        MESSAGE lx_error->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
    ENDTRY.

    " Warnings do not block the output - show them first
    DATA(lt_warnings) = lo_data->get_warnings( ).
    IF lt_warnings IS NOT INITIAL.
      zcl_sd_epack_output=>show_messages( it_messages = lt_warnings
                                          iv_title    = 'Warnings - output will be produced' ).
    ENDIF.

    TRY.
        DATA(lv_produced) = NEW zcl_sd_epack_output( )->send( is_data     = ls_print
                                                              iv_output   = output( )
                                                              iv_device   = p_dest
                                                              iv_copies   = p_copy
                                                              iv_pdf_path = CONV #( p_path ) ).
        IF lv_produced = abap_true.
          zcl_sd_epack_data=>mark_printed( lv_packno ).
        ENDIF.
      CATCH zcx_sd_epack INTO lx_error.
        MESSAGE lx_error->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD format.
    rv_format = COND #( WHEN p_fwith = abap_true THEN zif_sd_epack=>gc_format-with_np
                        WHEN p_fwout = abap_true THEN zif_sd_epack=>gc_format-without_np
                        ELSE zif_sd_epack=>gc_format-auto ).
  ENDMETHOD.


  METHOD output.
    rv_output = COND #( WHEN p_print = abap_true THEN zif_sd_epack=>gc_output-print
                        WHEN p_pdf   = abap_true THEN zif_sd_epack=>gc_output-pdf
                        ELSE zif_sd_epack=>gc_output-preview ).
  ENDMETHOD.


  METHOD pdf_file_dialog.

    DATA lv_file TYPE string.
    DATA lv_path TYPE string.
    DATA lv_full TYPE string.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING  default_extension = 'pdf'
                 default_file_name = |PackingList_{ p_packno ALPHA = OUT }.pdf|
                 file_filter       = 'PDF (*.pdf)|*.pdf'
      CHANGING   filename          = lv_file
                 path              = lv_path
                 fullpath          = lv_full
      EXCEPTIONS OTHERS            = 1 ).
    IF sy-subrc = 0 AND lv_full IS NOT INITIAL.
      cv_path = lv_full.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Event blocks
*&---------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
  " PDF path only when "Download PDF" is selected
  LOOP AT SCREEN.
    IF screen-group1 = 'PDF'.
      screen-active = COND #( WHEN p_pdf = abap_true THEN '1' ELSE '0' ).
      MODIFY SCREEN.
    ENDIF.
  ENDLOOP.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  lcl_report=>pdf_file_dialog( CHANGING cv_path = p_path ).

AT SELECTION-SCREEN.
  IF sy-ucomm = 'ONLI' AND s_vbeln[] IS INITIAL AND p_packno IS INITIAL.
    MESSAGE e003(zsd_epack).
  ENDIF.

START-OF-SELECTION.
  NEW lcl_report( )->run( ).

*&---------------------------------------------------------------------*
*& Text elements
*&   B01  Document selection
*&   B02  Output format
*&   B03  Output options
*& Selection texts
*&   S_VBELN  Tax Invoice
*&   P_PACKNO Packing List No.
*&   P_FAUTO  Automatic (notify party if maintained)
*&   P_FWITH  With notify party
*&   P_FWOUT  Without notify party
*&   P_PREV   Print preview
*&   P_PRINT  Print
*&   P_PDF    Download PDF
*&   P_DEST   Output device
*&   P_COPY   Number of copies
*&   P_PATH   PDF file
*&---------------------------------------------------------------------*
