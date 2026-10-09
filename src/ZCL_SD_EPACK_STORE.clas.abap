"! <p class="shorttext synchronized">Export Packing List - save data and log</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Last step of a run ("at last add the data in the Z table"):
"!   1. SAVE_PACKING_LISTS - the uploaded data of every packing list whose
"!      form was produced is written to the FS Z table(s)
"!      ZSD_EPACK_HDR / ZSD_EPACK_INV / ZSD_EPACK_DATA. An existing packing
"!      list is replaced and gets the next revision number (A22).
"!   2. WRITE_LOG - one row per packing list (successful or not) and one
"!      row for file-level errors in ZSD_EPACK_LOG.
"!   3. FINISH - one COMMIT WORK for data and log, then release the locks.
"! All database writes are array operations - no INSERT / UPDATE in loops.
CLASS zcl_sd_epack_store DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "! @parameter iv_source | ZIF_SD_EPACK=>GC_SOURCE-*
    "! @parameter iv_output | ZIF_SD_EPACK=>GC_OUTPUT-*
    METHODS constructor
      IMPORTING iv_source   TYPE char1
                iv_filename TYPE string OPTIONAL
                iv_output   TYPE char1.

    "! Save the packing lists with RESULT-OUTPUT = X (Excel source only);
    "! sets RESULT-SAVED and adds messages to the result.
    METHODS save_packing_lists
      IMPORTING it_packing_lists TYPE zif_sd_epack=>tt_packing_list
      CHANGING  ct_result        TYPE zif_sd_epack=>tt_result.

    "! Print stamp for reprinted packing lists (print / PDF only)
    METHODS stamp_reprint
      IMPORTING it_result TYPE zif_sd_epack=>tt_result.

    "! One log row per packing list + one for file-level messages
    METHODS write_log
      IMPORTING it_packing_lists TYPE zif_sd_epack=>tt_packing_list
                it_result        TYPE zif_sd_epack=>tt_result
                it_file_messages TYPE bapiret2_t
      RETURNING VALUE(rv_runid)  TYPE sysuuid_c32.

    "! COMMIT WORK and release the locks
    METHODS finish.

  PRIVATE SECTION.

    DATA mv_source   TYPE char1.
    DATA mv_filename TYPE string.
    DATA mv_output   TYPE char1.
    DATA mt_locked   TYPE SORTED TABLE OF zsd_packno WITH UNIQUE KEY table_line.

    CLASS-METHODS log_message
      IMPORTING it_messages       TYPE bapiret2_t
      RETURNING VALUE(rs_message) TYPE bapiret2.

ENDCLASS.


CLASS zcl_sd_epack_store IMPLEMENTATION.

  METHOD constructor.
    mv_source   = iv_source.
    mv_filename = iv_filename.
    mv_output   = iv_output.
  ENDMETHOD.


  METHOD save_packing_lists.

    DATA lr_packno TYPE RANGE OF zsd_packno.
    DATA lt_hdr    TYPE STANDARD TABLE OF zsd_epack_hdr  WITH EMPTY KEY.
    DATA lt_inv    TYPE STANDARD TABLE OF zsd_epack_inv  WITH EMPTY KEY.
    DATA lt_item   TYPE STANDARD TABLE OF zsd_epack_data WITH EMPTY KEY.

    IF mv_source <> zif_sd_epack=>gc_source-excel.
      RETURN.
    ENDIF.

    " --- Which packing lists may be saved -----------------------------------
    LOOP AT ct_result ASSIGNING FIELD-SYMBOL(<ls_result>) WHERE output = abap_true.

      " Authorization to create / change packing lists of the company code
      AUTHORITY-CHECK OBJECT zif_sd_epack=>gc_auth_object
        ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-create
        ID 'BUKRS' FIELD <ls_result>-bukrs.
      IF sy-subrc <> 0.
        APPEND zcl_sd_epack_rules=>msg( iv_type = 'W' iv_no = '012' iv_v1 = 'Company code'
                                        iv_v2 = <ls_result>-bukrs iv_v3 = zif_sd_epack=>gc_actvt-create )
               TO <ls_result>-messages.
        CONTINUE.
      ENDIF.

      " Lock against a parallel run for the same packing list
      CALL FUNCTION 'ENQUEUE_EZSD_EPACK'
        EXPORTING
          mode_zsd_epack_hdr = 'E'
          mandt              = sy-mandt
          zsd_packno         = <ls_result>-packno
        EXCEPTIONS
          foreign_lock       = 1
          system_failure     = 2
          OTHERS             = 3.
      IF sy-subrc <> 0.
        APPEND zcl_sd_epack_rules=>msg( iv_type = 'W' iv_no = '014' iv_v1 = <ls_result>-packno
                                        iv_v2 = sy-msgv1 ) TO <ls_result>-messages.
        CONTINUE.
      ENDIF.
      INSERT <ls_result>-packno INTO TABLE mt_locked.

      APPEND VALUE #( sign = 'I' option = 'EQ' low = <ls_result>-packno ) TO lr_packno.
    ENDLOOP.

    IF lr_packno IS INITIAL.
      RETURN.
    ENDIF.

    " --- Revision and creation data of packing lists saved before ----------
    SELECT zsd_packno, zsd_revno, zsd_ernam, zsd_erdat, zsd_erzet
      FROM zsd_epack_hdr
      WHERE zsd_packno IN @lr_packno
      INTO TABLE @DATA(lt_existing).
    SORT lt_existing BY zsd_packno.

    DATA(lv_printed) = xsdbool( mv_output = zif_sd_epack=>gc_output-print
                                OR mv_output = zif_sd_epack=>gc_output-pdf ).

    LOOP AT lr_packno INTO DATA(ls_packno).
      DATA(ls_pl)     = it_packing_lists[ packno = ls_packno-low ].
      DATA(ls_result) = ct_result[ packno = ls_packno-low ].
      DATA(ls_hdr)    = ls_pl-header.

      ls_hdr-zsd_bukrs    = ls_result-bukrs.
      ls_hdr-zsd_filename = mv_filename.

      READ TABLE lt_existing INTO DATA(ls_old) WITH KEY zsd_packno = ls_packno-low BINARY SEARCH.
      IF sy-subrc = 0.
        ls_hdr-zsd_revno = ls_old-zsd_revno + 1.
        ls_hdr-zsd_ernam = ls_old-zsd_ernam.
        ls_hdr-zsd_erdat = ls_old-zsd_erdat.
        ls_hdr-zsd_erzet = ls_old-zsd_erzet.
        ls_hdr-zsd_aenam = sy-uname.
        ls_hdr-zsd_aedat = sy-datum.
        ls_hdr-zsd_aezet = sy-uzeit.
      ELSE.
        ls_hdr-zsd_revno = 0.
        ls_hdr-zsd_ernam = sy-uname.
        ls_hdr-zsd_erdat = sy-datum.
        ls_hdr-zsd_erzet = sy-uzeit.
      ENDIF.
      IF lv_printed = abap_true.
        ls_hdr-zsd_prnam = sy-uname.
        ls_hdr-zsd_prdat = sy-datum.
        ls_hdr-zsd_przet = sy-uzeit.
      ENDIF.
      APPEND ls_hdr TO lt_hdr.

      LOOP AT ls_pl-invoices INTO DATA(lv_vbeln).
        APPEND VALUE #( zsd_packno = ls_packno-low vbeln = lv_vbeln ) TO lt_inv.
      ENDLOOP.

      " Line items exactly as uploaded (not the material-master defaults)
      APPEND LINES OF ls_pl-items TO lt_item.
    ENDLOOP.

    " --- Replace old versions, write new ones (array operations) ------------
    DELETE FROM zsd_epack_data WHERE zsd_packno IN @lr_packno.
    DELETE FROM zsd_epack_inv  WHERE zsd_packno IN @lr_packno.

    MODIFY zsd_epack_hdr FROM TABLE @lt_hdr.
    DATA(lv_ok) = xsdbool( sy-subrc = 0 ).

    INSERT zsd_epack_inv FROM TABLE @lt_inv.
    IF sy-subrc <> 0.
      lv_ok = abap_false.
    ENDIF.

    INSERT zsd_epack_data FROM TABLE @lt_item.
    IF sy-subrc <> 0.
      lv_ok = abap_false.
    ENDIF.

    LOOP AT lr_packno INTO ls_packno.
      ASSIGN ct_result[ packno = ls_packno-low ] TO <ls_result>.
      IF lv_ok = abap_true.
        <ls_result>-saved = abap_true.
        APPEND zcl_sd_epack_rules=>msg( iv_type = 'S' iv_no = '032' iv_v1 = ls_packno-low
                                        iv_v2 = lt_hdr[ zsd_packno = ls_packno-low ]-zsd_revno
                                        iv_v3 = <ls_result>-item_count iv_v4 = <ls_result>-inv_count )
               TO <ls_result>-messages.
      ELSE.
        APPEND zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '052' iv_v1 = ls_packno-low )
               TO <ls_result>-messages.
      ENDIF.
    ENDLOOP.

    IF lv_ok = abap_false.
      " Nothing half-written: undo all data changes of this run
      ROLLBACK WORK.
    ENDIF.

  ENDMETHOD.


  METHOD stamp_reprint.

    DATA lr_packno TYPE RANGE OF zsd_packno.

    IF mv_source <> zif_sd_epack=>gc_source-saved
       OR mv_output = zif_sd_epack=>gc_output-preview.
      RETURN.
    ENDIF.

    lr_packno = VALUE #( FOR ls_res IN it_result WHERE ( output = abap_true )
                         ( sign = 'I' option = 'EQ' low = ls_res-packno ) ).
    IF lr_packno IS INITIAL.
      RETURN.
    ENDIF.

    UPDATE zsd_epack_hdr
      SET zsd_prnam = @sy-uname,
          zsd_prdat = @sy-datum,
          zsd_przet = @sy-uzeit
      WHERE zsd_packno IN @lr_packno.

  ENDMETHOD.


  METHOD write_log.

    DATA lt_log TYPE STANDARD TABLE OF zsd_epack_log WITH EMPTY KEY.

    TRY.
        rv_runid = cl_system_uuid=>create_uuid_c32_static( ).
      CATCH cx_uuid_error.
        rv_runid = |{ sy-datum }{ sy-uzeit }{ sy-uname }|.
    ENDTRY.

    DATA(ls_base) = VALUE zsd_epack_log( runid    = rv_runid
                                         source   = mv_source
                                         filename = mv_filename
                                         output   = mv_output
                                         ernam    = sy-uname
                                         erdat    = sy-datum
                                         erzet    = sy-uzeit ).

    " File-level problems (wrong layout, rows without packing list number)
    IF it_file_messages IS NOT INITIAL.
      DATA(ls_msg) = log_message( it_file_messages ).
      DATA(ls_file) = ls_base.
      ls_file-status   = COND #( WHEN zcl_sd_epack_rules=>has_errors( it_file_messages )
                                 THEN zif_sd_epack=>gc_status-error
                                 ELSE zif_sd_epack=>gc_status-warning ).
      ls_file-msg_cnt  = lines( it_file_messages ).
      ls_file-msg_type = ls_msg-type.
      ls_file-message  = ls_msg-message.
      APPEND ls_file TO lt_log.
    ENDIF.

    LOOP AT it_result INTO DATA(ls_result).
      DATA(ls_log) = ls_base.
      ls_msg = log_message( ls_result-messages ).

      ls_log-zsd_packno = ls_result-packno.
      ls_log-formname   = ls_result-print-form.
      ls_log-status     = ls_result-status.
      ls_log-inv_cnt    = ls_result-inv_count.
      ls_log-item_cnt   = ls_result-item_count.
      ls_log-output_done = ls_result-output.
      ls_log-saved      = ls_result-saved.
      ls_log-msg_cnt    = lines( ls_result-messages ).
      ls_log-msg_type   = ls_msg-type.
      ls_log-message    = ls_msg-message.

      DATA(ls_pl) = VALUE zif_sd_epack=>ts_packing_list( it_packing_lists[ packno = ls_result-packno ] OPTIONAL ).
      ls_log-invoices = concat_lines_of( table = VALUE string_table( FOR lv_vbeln IN ls_pl-invoices
                                                                     ( |{ lv_vbeln ALPHA = OUT }| ) )
                                         sep   = `, ` ).
      ls_log-xls_rows = ls_pl-rows.

      APPEND ls_log TO lt_log.
    ENDLOOP.

    IF lt_log IS NOT INITIAL.
      INSERT zsd_epack_log FROM TABLE @lt_log.
      IF sy-subrc <> 0.
        MESSAGE s062(zsd_epack) DISPLAY LIKE 'E'.
      ENDIF.
    ENDIF.

  ENDMETHOD.


  METHOD finish.

    COMMIT WORK AND WAIT.

    LOOP AT mt_locked INTO DATA(lv_packno).
      CALL FUNCTION 'DEQUEUE_EZSD_EPACK'
        EXPORTING
          mode_zsd_epack_hdr = 'E'
          mandt              = sy-mandt
          zsd_packno         = lv_packno.
    ENDLOOP.
    CLEAR mt_locked.

  ENDMETHOD.


  METHOD log_message.

    " The most important message: first error, else first warning, else
    " the last success message (e.g. "saved")
    LOOP AT it_messages INTO rs_message WHERE type CA 'EAX'.
      RETURN.
    ENDLOOP.
    LOOP AT it_messages INTO rs_message WHERE type = 'W'.
      RETURN.
    ENDLOOP.
    rs_message = VALUE #( it_messages[ lines( it_messages ) ] OPTIONAL ).

  ENDMETHOD.

ENDCLASS.
