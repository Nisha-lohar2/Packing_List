"! <p class="shorttext synchronized">Export Packing List - form output</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Calls one of the two Smart Forms (FS: "with notifier party" /
"! "without notifier party") for preview, print or PDF download
"! (output channel not specified in the FS - ASSUMPTION A25) and shows
"! message lists in a popup.
CLASS zcl_sd_epack_output DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "! @parameter iv_output   | ZIF_SD_EPACK=>GC_OUTPUT-*
    "! @parameter rv_produced | X = form processed (preview shown, spool
    "!                          request created or PDF written);
    "!                          space = cancelled by the user
    "! @raising zcx_sd_epack  | form missing or form processing error
    METHODS send
      IMPORTING is_data            TYPE zif_sd_epack=>ts_print_data
                iv_output          TYPE char1
                iv_device          TYPE rspopname OPTIONAL
                iv_copies          TYPE tdcopies DEFAULT 1
                iv_pdf_path        TYPE string OPTIONAL
      RETURNING VALUE(rv_produced) TYPE abap_bool
      RAISING   zcx_sd_epack.

    "! Message list as ALV popup
    CLASS-METHODS show_messages
      IMPORTING it_messages TYPE bapiret2_t
                iv_title    TYPE lvc_title OPTIONAL.

    "! Result of a run: one block of lines per packing list (status, form,
    "! counts, output / saved flags and every message) + file messages
    CLASS-METHODS show_result
      IMPORTING it_result        TYPE zif_sd_epack=>tt_result
                it_file_messages TYPE bapiret2_t
                iv_runid         TYPE sysuuid_c32 OPTIONAL.

  PRIVATE SECTION.

    "! Default printer, read once per run (SEND is called once per packing list)
    DATA mv_default_device TYPE rspopname.

    METHODS download_pdf
      IMPORTING it_otf  TYPE tsfotf
                iv_path TYPE string
      RAISING   zcx_sd_epack.

ENDCLASS.


CLASS zcl_sd_epack_output IMPLEMENTATION.

  METHOD send.

    DATA lv_fm  TYPE rs38l_fnam.
    DATA ls_job TYPE ssfcrescl.

    CALL FUNCTION 'SSF_FUNCTION_MODULE_NAME'
      EXPORTING
        formname           = is_data-form
      IMPORTING
        fm_name            = lv_fm
      EXCEPTIONS
        no_form            = 1
        no_function_module = 2
        OTHERS             = 3.
    IF sy-subrc <> 0.
      zcx_sd_epack=>raise( iv_msgno = '040' iv_v1 = is_data-form ).
    ENDIF.

    " Output device: entered on the selection screen, else the user's
    " default printer, else LOCL
    DATA(lv_device) = iv_device.
    IF lv_device IS INITIAL.
      IF mv_default_device IS INITIAL.
        SELECT SINGLE spld FROM usr01 WHERE bname = @sy-uname INTO @mv_default_device.
        IF mv_default_device IS INITIAL.
          mv_default_device = 'LOCL'.
        ENDIF.
      ENDIF.
      lv_device = mv_default_device.
    ENDIF.

    DATA(ls_control) = VALUE ssfctrlop(
      no_dialog = abap_true
      preview   = xsdbool( iv_output = zif_sd_epack=>gc_output-preview )
      getotf    = xsdbool( iv_output = zif_sd_epack=>gc_output-pdf ) ).

    DATA(ls_options) = VALUE ssfcompop(
      tddest   = lv_device
      tdcopies = COND #( WHEN iv_copies IS INITIAL THEN 1 ELSE iv_copies )
      tdimmed  = xsdbool( iv_output = zif_sd_epack=>gc_output-print )
      tdnewid  = abap_true ).

    CALL FUNCTION lv_fm
      EXPORTING
        control_parameters = ls_control
        output_options     = ls_options
        user_settings      = space
        is_header          = is_data-header
        it_text            = is_data-texts
        it_item            = is_data-items
      IMPORTING
        job_output_info    = ls_job
      EXCEPTIONS
        formatting_error   = 1
        internal_error     = 2
        send_error         = 3
        user_canceled      = 4
        OTHERS             = 5.

    CASE sy-subrc.
      WHEN 0.
      WHEN 4.
        RETURN.                               " user cancelled
      WHEN OTHERS.
        zcx_sd_epack=>raise( iv_msgno = '039' iv_v1 = sy-subrc iv_v2 = is_data-form ).
    ENDCASE.

    IF iv_output = zif_sd_epack=>gc_output-pdf.
      download_pdf( it_otf  = ls_job-otfdata
                    iv_path = iv_pdf_path ).
    ENDIF.
    rv_produced = abap_true.

  ENDMETHOD.


  METHOD download_pdf.

    DATA lv_size  TYPE i.
    DATA lv_pdf   TYPE xstring.
    DATA lt_lines TYPE STANDARD TABLE OF tline WITH EMPTY KEY.
    DATA lt_otf   TYPE tsfotf.

    lt_otf = it_otf.

    CALL FUNCTION 'CONVERT_OTF'
      EXPORTING
        format                = 'PDF'
      IMPORTING
        bin_filesize          = lv_size
        bin_file              = lv_pdf
      TABLES
        otf                   = lt_otf
        lines                 = lt_lines
      EXCEPTIONS
        err_max_linewidth     = 1
        err_format            = 2
        err_conv_not_possible = 3
        err_bad_otf           = 4
        OTHERS                = 5.
    IF sy-subrc <> 0.
      zcx_sd_epack=>raise( iv_msgno = '039' iv_v1 = sy-subrc iv_v2 = 'CONVERT_OTF' ).
    ENDIF.

    DATA(lt_solix) = cl_bcs_convert=>xstring_to_solix( lv_pdf ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING  bin_filesize = lv_size
                 filename     = iv_path
                 filetype     = 'BIN'
      CHANGING   data_tab     = lt_solix
      EXCEPTIONS OTHERS       = 1 ).
    IF sy-subrc <> 0.
      zcx_sd_epack=>raise( iv_msgno = '053' iv_v1 = iv_path ).
    ENDIF.


  ENDMETHOD.


  METHOD show_messages.

    TYPES: BEGIN OF ty_display,
             icon    TYPE icon_d,
             type    TYPE bapi_mtype,
             number  TYPE symsgno,
             message TYPE bapi_msg,
           END OF ty_display.
    DATA lt_display TYPE STANDARD TABLE OF ty_display WITH EMPTY KEY.

    IF it_messages IS INITIAL.
      RETURN.
    ENDIF.

    lt_display = VALUE #( FOR ls_msg IN it_messages
                          ( icon    = SWITCH #( ls_msg-type
                                                WHEN 'E' OR 'A' OR 'X' THEN icon_led_red
                                                WHEN 'W'               THEN icon_led_yellow
                                                ELSE                        icon_led_green )
                            type    = ls_msg-type
                            number  = ls_msg-number
                            message = ls_msg-message ) ).

    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_alv)
                                CHANGING  t_table      = lt_display ).
        lo_alv->get_columns( )->set_optimize( abap_true ).
        lo_alv->get_display_settings( )->set_list_header(
          COND #( WHEN iv_title IS INITIAL THEN 'Export Packing List - messages' ELSE iv_title ) ).
        lo_alv->set_screen_popup( start_column = 5  end_column = 160
                                  start_line   = 3  end_line   = 25 ).
        lo_alv->display( ).
      CATCH cx_salv_msg.
        " Fallback: first message only
        DATA(ls_first) = it_messages[ 1 ].
        MESSAGE ls_first-message TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD show_result.

    TYPES: BEGIN OF ty_line,
             icon    TYPE icon_d,
             packno  TYPE zsd_packno,
             status  TYPE zsd_epack_status,
             form    TYPE tdsfname,
             inv     TYPE i,
             items   TYPE i,
             output  TYPE xfeld,
             saved   TYPE xfeld,
             type    TYPE bapi_mtype,
             number  TYPE symsgno,
             message TYPE bapi_msg,
           END OF ty_line.
    DATA lt_line TYPE STANDARD TABLE OF ty_line WITH EMPTY KEY.

    LOOP AT it_file_messages INTO DATA(ls_file).
      APPEND VALUE #( icon = icon_led_red type = ls_file-type number = ls_file-number
                      message = ls_file-message ) TO lt_line.
    ENDLOOP.

    LOOP AT it_result INTO DATA(ls_result).
      DATA(ls_base) = VALUE ty_line(
        icon   = SWITCH #( ls_result-status WHEN zif_sd_epack=>gc_status-error   THEN icon_led_red
                                            WHEN zif_sd_epack=>gc_status-warning THEN icon_led_yellow
                                            ELSE icon_led_green )
        packno = ls_result-packno
        status = ls_result-status
        form   = ls_result-print-form
        inv    = ls_result-inv_count
        items  = ls_result-item_count
        output = ls_result-output
        saved  = ls_result-saved ).
      IF ls_result-messages IS INITIAL.
        APPEND ls_base TO lt_line.
      ENDIF.
      LOOP AT ls_result-messages INTO DATA(ls_msg).
        DATA(ls_line) = ls_base.
        ls_line-type    = ls_msg-type.
        ls_line-number  = ls_msg-number.
        ls_line-message = ls_msg-message.
        APPEND ls_line TO lt_line.
      ENDLOOP.
    ENDLOOP.

    IF lt_line IS INITIAL.
      RETURN.
    ENDIF.

    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_alv)
                                CHANGING  t_table      = lt_line ).
        lo_alv->get_functions( )->set_all( abap_true ).
        lo_alv->get_columns( )->set_optimize( abap_true ).
        DATA(lo_columns) = lo_alv->get_columns( ).
        lo_columns->get_column( 'INV' )->set_short_text( 'Invoices' ).
        lo_columns->get_column( 'ITEMS' )->set_short_text( 'Items' ).
        lo_columns->get_column( 'OUTPUT' )->set_short_text( 'Output' ).
        lo_columns->get_column( 'SAVED' )->set_short_text( 'Saved' ).
        lo_alv->get_sorts( )->add_sort( columnname = 'PACKNO' ).
        lo_alv->get_display_settings( )->set_list_header(
          CONV #( |Export Packing List - { lines( it_result ) } packing list(s), run { iv_runid }| ) ).
        lo_alv->display( ).
      CATCH cx_salv_msg cx_salv_not_found cx_salv_existing cx_salv_data_error.
        MESSAGE s063(zsd_epack) DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.

ENDCLASS.
