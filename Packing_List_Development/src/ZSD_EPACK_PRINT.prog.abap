*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_PRINT
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list"
*& Astral Limited / Project UDAY / SD
*& Transaction code : ZSD_EPACK  (as named in the FS)
*&
*& 1. Upload the packing list Excel file (FS template "Z table
*&    Format1.xlsx") - it may contain several packing lists.
*& 2. For every packing list: read its invoices' header data and its line
*&    items exactly as described in the FS logic sheet (Logic_sheet1.xlsx).
*& 3. Output ONE Smart Form per packing list number (with / without
*&    notify party). An error in one packing list never stops the others.
*& 4. At the end, save the data of every packing list that was output to
*&    the Z tables and write one log row per packing list (ZSD_EPACK_LOG).
*&
*& Option "Saved packing lists" reprints packing lists saved earlier,
*& selected by tax invoice and/or packing list number.
*&
*& The program is only a shell. All logic is in:
*&   ZCL_SD_EPACK_UPLOAD  - Excel parsing, split by packing list (unit-tested)
*&   ZCL_SD_EPACK_DATA    - SAP data per logic sheet, bulk reads, checks
*&   ZCL_SD_EPACK_RULES   - weight rule, totals, formatting (unit-tested)
*&   ZCL_SD_EPACK_OUTPUT  - Smart Form call, PDF, result list
*&   ZCL_SD_EPACK_STORE   - save to Z tables, log
*&---------------------------------------------------------------------*
REPORT zsd_epack_print.

DATA gv_vbeln  TYPE vbeln_vf.
DATA gv_packno TYPE zsd_packno.

*&---------------------------------------------------------------------*
*& Selection screen
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_xls   RADIOBUTTON GROUP src DEFAULT 'X' USER-COMMAND src.
PARAMETERS p_file  TYPE rlgrap-filename LOWER CASE MODIF ID xls.
PARAMETERS p_saved RADIOBUTTON GROUP src.
SELECT-OPTIONS s_vbeln  FOR gv_vbeln  MODIF ID sav.
SELECT-OPTIONS s_packno FOR gv_packno MODIF ID sav.
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
PARAMETERS p_dir   TYPE rlgrap-filename LOWER CASE MODIF ID pdf.
PARAMETERS p_check AS CHECKBOX.
SELECTION-SCREEN END OF BLOCK b3.

*&---------------------------------------------------------------------*
*& Local controller
*&---------------------------------------------------------------------*
CLASS lcl_report DEFINITION FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS run.

    CLASS-METHODS file_dialog
      CHANGING cv_file TYPE rlgrap-filename.

    CLASS-METHODS folder_dialog
      CHANGING cv_dir TYPE rlgrap-filename.

  PRIVATE SECTION.
    DATA mt_packing_lists TYPE zif_sd_epack=>tt_packing_list.
    DATA mt_file_messages TYPE bapiret2_t.
    DATA mt_result        TYPE zif_sd_epack=>tt_result.

    "! Step 1 - packing lists from the Excel file or from the Z tables
    METHODS collect_packing_lists
      RETURNING VALUE(rv_ok) TYPE abap_bool.

    "! Step 3 - one form per packing list
    METHODS output_forms.

    "! Step 4 - save + log, then the result list
    METHODS finish.

    METHODS format
      RETURNING VALUE(rv_format) TYPE char1.

    METHODS output
      RETURNING VALUE(rv_output) TYPE char1.

    METHODS pdf_path
      IMPORTING iv_packno      TYPE zsd_packno
      RETURNING VALUE(rv_path) TYPE string.

ENDCLASS.


CLASS lcl_report IMPLEMENTATION.

  METHOD run.

    IF collect_packing_lists( ) = abap_false.
      finish( ).
      RETURN.
    ENDIF.

    " Step 2 - header and item data of ALL packing lists in one go
    mt_result = NEW zcl_sd_epack_data( )->prepare( it_packing_lists = mt_packing_lists
                                                    iv_format        = format( ) ).

    IF p_check = abap_false.
      output_forms( ).
    ENDIF.

    finish( ).

  ENDMETHOD.


  METHOD collect_packing_lists.

    IF p_xls = abap_true.
      TRY.
          DATA(lv_xstring) = zcl_sd_epack_upload=>read_file( CONV #( p_file ) ).
        CATCH zcx_sd_epack INTO DATA(lx_error).
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'E'
                                          iv_no   = lx_error->if_t100_message~t100key-msgno
                                          iv_v1   = lx_error->mv_msgv1
                                          iv_v2   = lx_error->mv_msgv2 ) TO mt_file_messages.
          RETURN.
      ENDTRY.

      NEW zcl_sd_epack_upload( )->parse( EXPORTING iv_xstring       = lv_xstring
                                                   iv_filename      = CONV #( p_file )
                                         IMPORTING et_packing_lists = mt_packing_lists
                                                   et_messages      = mt_file_messages ).
    ELSE.
      TRY.
          NEW zcl_sd_epack_data( )->load_saved( EXPORTING it_vbeln         = s_vbeln[]
                                                          it_packno        = s_packno[]
                                                IMPORTING et_packing_lists = mt_packing_lists
                                                          et_messages      = mt_file_messages ).
        CATCH zcx_sd_epack INTO lx_error.
          mt_file_messages = lx_error->mt_messages.
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'E'
                                          iv_no   = lx_error->if_t100_message~t100key-msgno
                                          iv_v1   = lx_error->mv_msgv1 ) TO mt_file_messages.
          RETURN.
      ENDTRY.
    ENDIF.

    rv_ok = xsdbool( mt_packing_lists IS NOT INITIAL ).

  ENDMETHOD.


  METHOD output_forms.

    DATA(lo_output) = NEW zcl_sd_epack_output( ).

    " One form per packing list number - the result table is keyed by
    " packing list, and every entry carries only its own header and items
    LOOP AT mt_result ASSIGNING FIELD-SYMBOL(<ls_result>)
         WHERE status <> zif_sd_epack=>gc_status-error.
      TRY.
          <ls_result>-output = lo_output->send( is_data     = <ls_result>-print
                                                iv_output   = output( )
                                                iv_device   = p_dest
                                                iv_copies   = p_copy
                                                iv_pdf_path = pdf_path( <ls_result>-packno ) ).
          IF <ls_result>-output = abap_true.
            APPEND zcl_sd_epack_rules=>msg( iv_type = 'S' iv_no = '065'
                                            iv_v1 = <ls_result>-print-form
                                            iv_v2 = <ls_result>-packno ) TO <ls_result>-messages.
            IF p_pdf = abap_true.
              APPEND zcl_sd_epack_rules=>msg( iv_type = 'S' iv_no = '038'
                                              iv_v1 = pdf_path( <ls_result>-packno ) ) TO <ls_result>-messages.
            ENDIF.
          ELSE.
            APPEND zcl_sd_epack_rules=>msg( iv_type = 'W' iv_no = '064'
                                            iv_v1 = <ls_result>-packno ) TO <ls_result>-messages.
            <ls_result>-status = zif_sd_epack=>gc_status-warning.
          ENDIF.

        CATCH zcx_sd_epack INTO DATA(lx_error).
          " Logged for this packing list; the loop continues with the next
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'E'
                                          iv_no   = lx_error->if_t100_message~t100key-msgno
                                          iv_v1   = lx_error->mv_msgv1
                                          iv_v2   = lx_error->mv_msgv2 ) TO <ls_result>-messages.
          <ls_result>-status = zif_sd_epack=>gc_status-error.
          <ls_result>-output = abap_false.
      ENDTRY.
    ENDLOOP.

  ENDMETHOD.


  METHOD finish.

    DATA lv_runid TYPE sysuuid_c32.

    " "Validate only" does not change the database
    IF p_check = abap_false.
      DATA(lo_store) = NEW zcl_sd_epack_store(
                             iv_source   = COND #( WHEN p_xls = abap_true THEN zif_sd_epack=>gc_source-excel
                                                                          ELSE zif_sd_epack=>gc_source-saved )
                             iv_filename = CONV #( p_file )
                             iv_output   = output( ) ).
      lo_store->save_packing_lists( EXPORTING it_packing_lists = mt_packing_lists
                                    CHANGING  ct_result        = mt_result ).
      lo_store->stamp_reprint( mt_result ).
      lv_runid = lo_store->write_log( it_packing_lists = mt_packing_lists
                                      it_result        = mt_result
                                      it_file_messages = mt_file_messages ).
      lo_store->finish( ).
    ENDIF.

    " Overall message
    DATA(lv_total)  = lines( mt_result ).
    DATA(lv_output) = REDUCE i( INIT n = 0 FOR r IN mt_result WHERE ( output = abap_true ) NEXT n = n + 1 ).
    DATA(lv_saved)  = REDUCE i( INIT n = 0 FOR r IN mt_result WHERE ( saved = abap_true ) NEXT n = n + 1 ).
    DATA(lv_error)  = REDUCE i( INIT n = 0 FOR r IN mt_result
                                WHERE ( status = zif_sd_epack=>gc_status-error ) NEXT n = n + 1 ).

    IF lv_total = 0.
      MESSAGE s059(zsd_epack) DISPLAY LIKE 'E'.
    ELSEIF p_check = abap_true.
      MESSAGE s067(zsd_epack) WITH lv_total lv_error.
    ELSEIF lv_error > 0.
      MESSAGE s066(zsd_epack) WITH lv_output lv_total lv_saved lv_error DISPLAY LIKE 'W'.
    ELSE.
      MESSAGE s066(zsd_epack) WITH lv_output lv_total lv_saved lv_error.
    ENDIF.

    zcl_sd_epack_output=>show_result( it_result        = mt_result
                                      it_file_messages = mt_file_messages
                                      iv_runid         = lv_runid ).

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


  METHOD pdf_path.

    DATA lv_sep TYPE c LENGTH 1.

    IF p_pdf = abap_false.
      RETURN.
    ENDIF.

    cl_gui_frontend_services=>get_file_separator( CHANGING   file_separator = lv_sep
                                                  EXCEPTIONS OTHERS         = 1 ).
    IF sy-subrc <> 0 OR lv_sep IS INITIAL.
      lv_sep = '\'.
    ENDIF.

    " One PDF file per packing list number
    rv_path = |{ condense( CONV string( p_dir ) ) }{ lv_sep }PackingList_{ iv_packno ALPHA = OUT }.pdf|.

  ENDMETHOD.


  METHOD file_dialog.

    DATA lt_files TYPE filetable.
    DATA lv_rc    TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING  file_filter    = 'Excel (*.xlsx)|*.xlsx'
                 multiselection = abap_false
      CHANGING   file_table     = lt_files
                 rc             = lv_rc
      EXCEPTIONS OTHERS         = 1 ).
    IF sy-subrc = 0 AND lv_rc = 1.
      cv_file = lt_files[ 1 ]-filename.
    ENDIF.

  ENDMETHOD.


  METHOD folder_dialog.

    DATA lv_folder TYPE string.

    cl_gui_frontend_services=>directory_browse(
      CHANGING   selected_folder = lv_folder
      EXCEPTIONS OTHERS          = 1 ).
    IF sy-subrc = 0 AND lv_folder IS NOT INITIAL.
      cv_dir = lv_folder.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Event blocks
*&---------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.
    CASE screen-group1.
      WHEN 'XLS'.
        screen-active = COND #( WHEN p_xls = abap_true THEN '1' ELSE '0' ).
      WHEN 'SAV'.
        screen-active = COND #( WHEN p_saved = abap_true THEN '1' ELSE '0' ).
      WHEN 'PDF'.
        screen-active = COND #( WHEN p_pdf = abap_true THEN '1' ELSE '0' ).
      WHEN OTHERS.
        CONTINUE.
    ENDCASE.
    MODIFY SCREEN.
  ENDLOOP.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_report=>file_dialog( CHANGING cv_file = p_file ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_dir.
  lcl_report=>folder_dialog( CHANGING cv_dir = p_dir ).

AT SELECTION-SCREEN.
  IF sy-ucomm = 'ONLI' OR sy-ucomm = 'PRIN'.
    IF p_xls = abap_true AND p_file IS INITIAL.
      MESSAGE e017(zsd_epack) WITH '-' 'Excel file'.
    ENDIF.
    IF p_saved = abap_true AND s_vbeln[] IS INITIAL AND s_packno[] IS INITIAL.
      MESSAGE e003(zsd_epack).
    ENDIF.
    IF p_pdf = abap_true AND p_check = abap_false AND p_dir IS INITIAL.
      MESSAGE e054(zsd_epack).
    ENDIF.
  ENDIF.

START-OF-SELECTION.
  NEW lcl_report( )->run( ).

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
