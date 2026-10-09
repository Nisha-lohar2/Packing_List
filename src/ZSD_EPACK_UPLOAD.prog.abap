*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_UPLOAD
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list"
*& Astral Limited / Project UDAY / SD
*& Transaction code : ZSD_EPACK_UPL  (not named in the FS - ASSUMPTION A23)
*&
*& FS 1: "Any information that is not available in SAP should be
*& maintained by users through an Excel upload into a custom Z-table."
*&
*& Uploads the FS Excel template (Z table Format1.xlsx layout) into
*& ZSD_EPACK_HDR / ZSD_EPACK_INV / ZSD_EPACK_DATA. A test run checks the
*& file without saving. All logic lives in ZCL_SD_EPACK_UPLOAD.
*&---------------------------------------------------------------------*
REPORT zsd_epack_upload.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_file TYPE rlgrap-filename LOWER CASE MODIF ID fil.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-b02.
PARAMETERS p_create RADIOBUTTON GROUP mod DEFAULT 'X' USER-COMMAND mod.
PARAMETERS p_overwr RADIOBUTTON GROUP mod.
PARAMETERS p_delete RADIOBUTTON GROUP mod.
PARAMETERS p_packno TYPE zsd_packno MODIF ID del.
PARAMETERS p_test   AS CHECKBOX DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK b2.

*&---------------------------------------------------------------------*
*& Local controller
*&---------------------------------------------------------------------*
CLASS lcl_report DEFINITION FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS run.

    CLASS-METHODS file_dialog
      CHANGING cv_file TYPE rlgrap-filename.

ENDCLASS.


CLASS lcl_report IMPLEMENTATION.

  METHOD run.

    DATA lt_messages TYPE bapiret2_t.

    IF p_delete = abap_true.
      IF p_packno IS INITIAL.
        MESSAGE s017(zsd_epack) WITH '-' 'Packing List No.' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      lt_messages = zcl_sd_epack_upload=>delete( iv_packno = p_packno
                                                 iv_test   = p_test ).
    ELSE.
      IF p_file IS INITIAL.
        MESSAGE s017(zsd_epack) WITH '-' 'Excel file' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      DATA(lo_upload) = NEW zcl_sd_epack_upload(
                              iv_mode = COND #( WHEN p_overwr = abap_true
                                                THEN zif_sd_epack=>gc_upload_mode-overwrite
                                                ELSE zif_sd_epack=>gc_upload_mode-create )
                              iv_test = p_test ).
      lt_messages = lo_upload->upload_file( CONV #( p_file ) ).
    ENDIF.

    zcl_sd_epack_output=>show_messages(
      it_messages = lt_messages
      iv_title    = COND #( WHEN p_test = abap_true THEN 'Packing list upload - TEST RUN'
                            ELSE 'Packing list upload' ) ).

  ENDMETHOD.


  METHOD file_dialog.

    DATA lt_files TYPE filetable.
    DATA lv_rc    TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING  file_filter = 'Excel (*.xlsx)|*.xlsx'
                 multiselection = abap_false
      CHANGING   file_table  = lt_files
                 rc          = lv_rc
      EXCEPTIONS OTHERS      = 1 ).
    IF sy-subrc = 0 AND lv_rc = 1.
      cv_file = lt_files[ 1 ]-filename.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Event blocks
*&---------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.
    CASE screen-group1.
      WHEN 'FIL'.
        screen-active = COND #( WHEN p_delete = abap_true THEN '0' ELSE '1' ).
      WHEN 'DEL'.
        screen-active = COND #( WHEN p_delete = abap_true THEN '1' ELSE '0' ).
      WHEN OTHERS.
        CONTINUE.
    ENDCASE.
    MODIFY SCREEN.
  ENDLOOP.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_report=>file_dialog( CHANGING cv_file = p_file ).

START-OF-SELECTION.
  NEW lcl_report( )->run( ).

*&---------------------------------------------------------------------*
*& Text elements
*&   B01  Excel file (template "Z table Format1.xlsx")
*&   B02  Processing
*& Selection texts
*&   P_FILE   Excel file
*&   P_CREATE Create new packing list
*&   P_OVERWR Overwrite existing packing list (new revision)
*&   P_DELETE Delete packing list
*&   P_PACKNO Packing List No.
*&   P_TEST   Test run (check only, no update)
*&---------------------------------------------------------------------*
