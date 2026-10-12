*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_F01
*&---------------------------------------------------------------------*
*& Main routines:
*&   1. Excel upload -> packing lists        (FORM parse_*, map_*)
*&   2. SAP data per FS logic sheet, bulk     (FORM read_sap_data, build_one)
*&   3. One Adobe Form per packing list       (FORM output_forms)
*&   4. Save to Z table ZSD_EPACK_DATA (log) (FORM save_packing_lists)
*&
*& PERFORMANCE: READ_SAP_DATA reads every SAP table ONCE for the invoices
*& of ALL packing lists (FOR ALL ENTRIES on de-duplicated keys). Each
*& packing list is then built from internal HASHED / SORTED tables only.
*& There is no SELECT inside any loop.
*&
*& WHERE THE LOGIC SHEET NAMES A FIELD THAT DOES NOT EXIST IN THE NAMED
*& TABLE, the table that holds the field is read instead:
*&   B4  BEZEI is in T005U (not T005S)   B5  LANDX is in T005T (not T005)
*&   B54 MAKTX is in MAKT (not MARA)     B13 AUBEL is a VBRP field
*&   B20/B26 "PARVW = SP / SH" = internal AG / WE; partner = KNVP-KUNN2
*& ⚠ SAP-VERIFY before activation: J_1BBRANCH-GSTIN, T001Z, T604N
*&   (SPRAS/LAND1/STEUC/TEXT1), TVZBT-ZTAGG, KNVP-DEFPA.
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& Main control
*&---------------------------------------------------------------------*
FORM main.

  DATA lv_ok TYPE abap_bool.

  " Step 1 - packing lists from the Excel file
  PERFORM collect_packing_lists CHANGING lv_ok.

  IF lv_ok = abap_true.
    " Step 2 - header and item data of ALL packing lists in one go
    PERFORM read_sap_data USING gt_packing_lists.
    PERFORM prepare_all.

    " Step 3 - one form per packing list
    PERFORM output_forms.
  ENDIF.

  " Step 4 - write the data to the log table ZSD_EPACK_DATA, then the
  " result list
  PERFORM save_packing_lists.
  PERFORM finish_db.

  PERFORM final_message.
  PERFORM show_result.

ENDFORM.

*&---------------------------------------------------------------------*
*& Step 1: packing lists from the Excel file
*&---------------------------------------------------------------------*
FORM collect_packing_lists CHANGING cv_ok TYPE abap_bool.

  DATA lv_xstring TYPE xstring.
  DATA lt_rows    TYPE ty_row_tab.
  DATA lv_errors  TYPE abap_bool.

  cv_ok = abap_false.

  PERFORM read_file USING p_file CHANGING lv_xstring gt_file_messages.
  PERFORM has_errors USING gt_file_messages CHANGING lv_errors.
  IF lv_errors = abap_true.
    RETURN.
  ENDIF.

  PERFORM read_worksheet USING lv_xstring p_file CHANGING lt_rows gt_file_messages.
  PERFORM has_errors USING gt_file_messages CHANGING lv_errors.
  IF lv_errors = abap_true.
    RETURN.
  ENDIF.

  PERFORM parse_rows USING lt_rows p_file CHANGING gt_packing_lists gt_file_messages.

  cv_ok = xsdbool( gt_packing_lists IS NOT INITIAL ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Read the Excel file from the presentation server
*&---------------------------------------------------------------------*
FORM read_file USING    iv_path     TYPE csequence
               CHANGING cv_xstring  TYPE xstring
                        ct_messages TYPE bapiret2_t.

  DATA lt_solix  TYPE solix_tab.
  DATA lv_length TYPE i.

  cl_gui_frontend_services=>gui_upload(
    EXPORTING  filename   = CONV #( iv_path )
               filetype   = 'BIN'
    IMPORTING  filelength = lv_length
    CHANGING   data_tab   = lt_solix
    EXCEPTIONS OTHERS     = 1 ).
  IF sy-subrc <> 0.
    PERFORM add_msg USING 'E' '015' iv_path 'GUI_UPLOAD' space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  IF lv_length = 0.
    PERFORM add_msg USING 'E' '036' space space space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  cv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_solix
                                                 iv_size  = lv_length ).

ENDFORM.

*&---------------------------------------------------------------------*
*& First worksheet -> rows of cell texts
*& CL_FDT_XL_SPREADSHEET reads .xlsx without OLE automation.
*& ⚠ SAP-VERIFY: not a released API; XCO_CP_XLSX on newer releases (A21).
*&---------------------------------------------------------------------*
FORM read_worksheet USING    iv_xstring  TYPE xstring
                             iv_filename TYPE csequence
                    CHANGING ct_rows     TYPE ty_row_tab
                             ct_messages TYPE bapiret2_t.

  DATA lo_excel  TYPE REF TO cl_fdt_xl_spreadsheet.
  DATA lr_sheet  TYPE REF TO data.
  DATA lx_excel  TYPE REF TO cx_fdt_excel_core.
  DATA ls_row    TYPE ty_row.
  DATA lv_error  TYPE string.
  FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

  CLEAR ct_rows.

  TRY.
      CREATE OBJECT lo_excel
        EXPORTING
          document_name = CONV string( iv_filename )
          xdocument     = iv_xstring.
      lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = DATA(lt_sheets) ).
      IF lt_sheets IS INITIAL.
        PERFORM add_msg USING 'E' '036' space space space space CHANGING ct_messages.
        RETURN.
      ENDIF.
      lr_sheet = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lt_sheets[ 1 ] ).
    CATCH cx_fdt_excel_core INTO lx_excel.
      lv_error = lx_excel->get_text( ).
      PERFORM add_msg USING 'E' '015' iv_filename lv_error space space CHANGING ct_messages.
      RETURN.
  ENDTRY.

  ASSIGN lr_sheet->* TO <lt_sheet>.
  IF sy-subrc <> 0.
    RETURN.
  ENDIF.

  LOOP AT <lt_sheet> ASSIGNING FIELD-SYMBOL(<ls_line>).
    CLEAR ls_row.
    ls_row-row = sy-tabix.
    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE <ls_line> TO FIELD-SYMBOL(<lv_cell>).
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      APPEND condense( val = CONV string( <lv_cell> ) del = ` ` ) TO ls_row-cells.
    ENDDO.
    " Skip completely empty rows
    LOOP AT ls_row-cells TRANSPORTING NO FIELDS WHERE table_line IS NOT INITIAL.
      APPEND ls_row TO ct_rows.
      EXIT.
    ENDLOOP.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Worksheet rows (row 1 = headings) -> packing lists
*& Several packing lists per file (A33): a row with "Packing list no"
*& filled starts / continues that packing list, rows with it empty belong
*& to the packing list above. Each packing list gets its own header,
*& invoices, items and messages.
*&---------------------------------------------------------------------*
FORM parse_rows USING    it_rows          TYPE ty_row_tab
                         iv_filename      TYPE csequence
                CHANGING ct_packing_lists TYPE ty_packing_list_tab
                         ct_messages      TYPE bapiret2_t.

  DATA lt_split  TYPE ty_pl_rows_tab.
  DATA ls_pl     TYPE ty_packing_list.
  DATA ls_head   TYPE ty_row.
  DATA lv_errors TYPE abap_bool.
  DATA lv_ext    TYPE string.

  TYPES: BEGIN OF ty_inv_pl,
           vbeln  TYPE vbeln_vf,
           packno TYPE zsd_packno,
         END OF ty_inv_pl.
  DATA lt_inv_pl TYPE SORTED TABLE OF ty_inv_pl WITH NON-UNIQUE KEY vbeln.

  CLEAR ct_packing_lists.

  IF lines( it_rows ) < 2.
    PERFORM add_msg USING 'E' '036' space space space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  READ TABLE it_rows INTO ls_head INDEX 1.
  PERFORM build_column_map USING ls_head CHANGING ct_messages.
  PERFORM has_errors USING ct_messages CHANGING lv_errors.
  IF lv_errors = abap_true.
    RETURN.
  ENDIF.

  DATA(lt_data_rows) = it_rows.
  DELETE lt_data_rows INDEX 1.

  PERFORM split_by_packing_list USING lt_data_rows CHANGING lt_split ct_messages.
  IF lt_split IS INITIAL.
    PERFORM add_msg USING 'E' '059' space space space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  " One packing list at a time: header, invoices, items, checks
  LOOP AT lt_split INTO DATA(ls_split).
    CLEAR ls_pl.
    ls_pl-packno              = ls_split-packno.
    ls_pl-header-zsd_packno   = ls_split-packno.
    ls_pl-rows = |{ ls_split-rows[ 1 ]-row }-{ ls_split-rows[ lines( ls_split-rows ) ]-row }|.

    PERFORM map_header USING ls_split-rows CHANGING ls_pl.
    PERFORM map_lines  USING ls_split-rows CHANGING ls_pl.

    " FS 2.1 weight rule and the quantity plausibility
    PERFORM check_weights    USING ls_pl-items CHANGING ls_pl-messages.
    PERFORM check_quantities USING ls_pl-items CHANGING ls_pl-messages.

    INSERT ls_pl INTO TABLE ct_packing_lists.
  ENDLOOP.

  " One invoice may not appear in two packing lists (A01)
  LOOP AT ct_packing_lists INTO ls_pl.
    LOOP AT ls_pl-invoices INTO DATA(lv_vbeln).
      INSERT VALUE #( vbeln = lv_vbeln packno = ls_pl-packno ) INTO TABLE lt_inv_pl.
    ENDLOOP.
  ENDLOOP.

  LOOP AT lt_inv_pl INTO DATA(ls_inv_pl).
    LOOP AT lt_inv_pl INTO DATA(ls_other) WHERE vbeln = ls_inv_pl-vbeln AND packno <> ls_inv_pl-packno.
      READ TABLE ct_packing_lists ASSIGNING FIELD-SYMBOL(<ls_pl>) WITH TABLE KEY packno = ls_inv_pl-packno.
      IF sy-subrc = 0.
        lv_ext = |{ ls_inv_pl-vbeln ALPHA = OUT }|.
        PERFORM add_msg USING 'E' '058' lv_ext ls_other-packno space space CHANGING <ls_pl>-messages.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Find every column by its heading (column order is free - A41).
*& FS template "Format": items A-K, invoice L, packing list M;
*& Sample_data_ABAP.xlsx: packing list A, invoice B, items C-M.
*&---------------------------------------------------------------------*
FORM build_column_map USING    is_row      TYPE ty_row
                      CHANGING ct_messages TYPE bapiret2_t.

  TYPES: BEGIN OF ty_key,
           name  TYPE string,
           regex TYPE string,
           label TYPE string,
           must  TYPE abap_bool,
         END OF ty_key.
  DATA lt_keys    TYPE STANDARD TABLE OF ty_key WITH EMPTY KEY.
  DATA lt_heading TYPE ty_text_tab.
  DATA lv_norm    TYPE string.
  DATA lv_col     TYPE i.

  CLEAR gs_col.

  lt_keys = VALUE #(
    ( name = `SN`       regex = `^(SLNO|SNO|SERIALNO)$`              label = `Sl.No.`             must = abap_true )
    ( name = `INV_SN`   regex = `^SRNOINV`                           label = `Sr.No. INV` )
    ( name = `ARTNO`    regex = `^ARTICLENO`                         label = `Article No.` )
    ( name = `STD`      regex = `^STD$`                              label = `STD` )
    ( name = `ART`      regex = `^(ARTICLES|NOOFARTICLES)$`          label = `Articles` )
    ( name = `PARTNO`   regex = `^PARTNO`                            label = `Part No.`           must = abap_true )
    ( name = `MATDESC`  regex = `^(PRODUCT|MATERIAL)DESCRIPTION`     label = `Product Description` )
    ( name = `QTY`      regex = `^(QUANTITY|QTY)`                    label = `Quantity`           must = abap_true )
    ( name = `GWT`      regex = `^GROSSW`                            label = `Gross Wt.` )
    ( name = `NWT`      regex = `^NETW`                              label = `Net Wt.` )
    ( name = `HSN`      regex = `^(HSCODE|HSNCODE|HSN)$`             label = `H.S Code` )
    ( name = `INVOICE`  regex = `^(SAP)?(TAX)?INVOICE(NUMBER|NO)S?$` label = `SAP Invoice Number` must = abap_true )
    ( name = `PACKNO`   regex = `^PA[A-Z]*LISTN(O|UMBER)$`           label = `Packing list no`    must = abap_true )
    ( name = `PKG_TYPE` regex = `^PACKAGETYPE`                       label = `Package Type` ) ).

  LOOP AT is_row-cells INTO DATA(lv_cell).
    PERFORM normalize_heading USING lv_cell CHANGING lv_norm.
    APPEND lv_norm TO lt_heading.
  ENDLOOP.

  LOOP AT lt_keys INTO DATA(ls_key).
    lv_col = 0.
    LOOP AT lt_heading INTO DATA(lv_heading).
      IF lv_heading IS NOT INITIAL AND matches( val = lv_heading regex = ls_key-regex ).
        lv_col = sy-tabix.
        EXIT.
      ENDIF.
    ENDLOOP.
    IF lv_col = 0 AND ls_key-must = abap_true.
      PERFORM add_msg USING 'E' '016' ls_key-label space space space CHANGING ct_messages.
    ENDIF.
    ASSIGN COMPONENT ls_key-name OF STRUCTURE gs_col TO FIELD-SYMBOL(<lv_col>).
    IF sy-subrc = 0.
      <lv_col> = lv_col.
    ENDIF.
  ENDLOOP.

  " Header columns (logic sheet "Pass ZSD_PACKNO get ...").
  " "Exporter ReF" is NOT mapped: Technical detail marks ZSD_EXPREF "Not
  " Req - fetched from SAP" (logic B11). Multi-line fields are wider than
  " in the FS because the sample outputs do not fit (A19).
  gt_map = VALUE #(
    ( regex = `^PA[A-Z]*LISTDATE$`        field = 'ZSD_PACKDT'    kind = 'D'               label = `Packing list date` )
    ( regex = `^BUYERSREFERENCENO`        field = 'ZSD_PARTY_REF' kind = 'C' maxlen = 200  label = `Buyer's Reference No.` )
    ( regex = `^BUYERSREFERENCEDATE`      field = 'ZSD_REF_DT'    kind = 'D'               label = `Buyer's Reference date` )
    ( regex = `^COUNTRYOFORIGIN`          field = 'ZSD_CNTY_ORGN' kind = 'C' maxlen = 20   label = `Country of origin` )
    ( regex = `^NOTIF(Y|IER)PARTY1$`      field = 'ZSD_NP1'       kind = 'T' maxlen = 1333 label = `Notifier Party1` )
    ( regex = `^NOTIF(Y|IER)PARTY2$`      field = 'ZSD_NP2'       kind = 'T' maxlen = 1333 label = `Notifier Party2` )
    ( regex = `^NOTIF(Y|IER)PARTY3$`      field = 'ZSD_NP3'       kind = 'T' maxlen = 1333 label = `Notifier Party3` )
    ( regex = `^NOTIF(Y|IER)PARTY4$`      field = 'ZSD_NP4'       kind = 'T' maxlen = 1333 label = `Notifier Party4` )
    ( regex = `^NOTIF(Y|IER)PARTY5$`      field = 'ZSD_NP5'       kind = 'T' maxlen = 1333 label = `Notifier Party5` )
    ( regex = `^PRECARRIAGE`              field = 'ZSD_PREC'      kind = 'C' maxlen = 30   label = `Pre-Carriage` )
    ( regex = `^PLACEOFRECEIPT`           field = 'ZSD_REC_PC'    kind = 'C' maxlen = 30   label = `Place of Receipt by Pre-carrier` )
    ( regex = `^PORTOFLOADING`            field = 'ZSD_POL'       kind = 'C' maxlen = 30   label = `Port of Loading` )
    ( regex = `^PORTOFDISCHARGE`          field = 'ZSD_POD'       kind = 'C' maxlen = 30   label = `Port of Discharge` )
    ( regex = `^PLACEOFDELIVERY`          field = 'ZSD_PLD'       kind = 'C' maxlen = 30   label = `Place of Delivery` )
    ( regex = `^MARKS`                    field = 'ZSD_MARK'      kind = 'T' maxlen = 255  label = `Marks & Nos.` )
    ( regex = `^VESSEL`                   field = 'ZSD_VSL_FLT'   kind = 'C' maxlen = 30   label = `Vessel / Flight No.` )
    ( regex = `^CONTAINERNO`              field = 'ZSD_CONT'      kind = 'T' maxlen = 1333 label = `Container No` )
    ( regex = `^NOOFPACKAGES`             field = 'ZSD_NOPACK'    kind = 'T' maxlen = 255  label = `No. of Packages` )
    ( regex = `^DESCRIPTIONOFGOODS`       field = 'ZSD_HSN_DESC'  kind = 'T' maxlen = 1333 label = `Description of goods` )
    ( regex = `^DECLARATIONOFPACKINGLIST` field = 'ZSD_PL_DECL'   kind = 'T' maxlen = 0    label = `Declaration of Packing list` )
    ( regex = `^ADVANCEAMOUNT`            field = 'ZSD_ADV_AMT'   kind = 'A'               label = `Advance amount received` )
    ( regex = `^COMMERCIALTERMS`          field = 'ZSD_CMMT'      kind = 'T' maxlen = 1333 label = `Commercial terms of invoice` )
    ( regex = `^DECLARATIONOFINVOICE`     field = 'ZSD_INV_DECL'  kind = 'T' maxlen = 0    label = `Declaration of Invoice` ) ).

  LOOP AT gt_map ASSIGNING FIELD-SYMBOL(<ls_map>).
    LOOP AT lt_heading INTO lv_heading.
      IF lv_heading IS NOT INITIAL AND matches( val = lv_heading regex = <ls_map>-regex ).
        <ls_map>-col = sy-tabix.
        EXIT.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

  READ TABLE gt_map INTO DATA(ls_date) WITH KEY field = 'ZSD_PACKDT'.
  IF sy-subrc <> 0 OR ls_date-col = 0.
    PERFORM add_msg USING 'E' '016' 'Packing list date' space space space CHANGING ct_messages.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Cell text; placeholders "Fetch from ..." count as empty (A41)
*&---------------------------------------------------------------------*
FORM get_cell USING    is_row  TYPE ty_row
                       iv_col  TYPE i
              CHANGING cv_text TYPE string.

  CLEAR cv_text.
  IF iv_col <= 0.
    RETURN.
  ENDIF.
  READ TABLE is_row-cells INTO cv_text INDEX iv_col.
  IF cv_text IS NOT INITIAL AND matches( val = to_upper( cv_text ) regex = `^\s*FETCH FROM.*` ).
    CLEAR cv_text.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Cell -> number with validation (message 018)
*&---------------------------------------------------------------------*
FORM get_number USING    is_row      TYPE ty_row
                         iv_col      TYPE i
                         iv_integer  TYPE abap_bool
                CHANGING cv_value    TYPE decfloat34
                         ct_messages TYPE bapiret2_t.

  DATA lv_text TYPE string.
  DATA lv_ok   TYPE abap_bool.
  DATA lv_col  TYPE string.

  CLEAR cv_value.
  PERFORM get_cell USING is_row iv_col CHANGING lv_text.
  IF lv_text IS INITIAL.
    RETURN.
  ENDIF.

  PERFORM parse_number USING lv_text CHANGING cv_value lv_ok.
  IF lv_ok = abap_false OR cv_value < 0
     OR ( iv_integer = abap_true AND ( frac( cv_value ) <> 0 OR cv_value > 999999 ) ).
    PERFORM column_name USING iv_col CHANGING lv_col.
    PERFORM add_msg USING 'E' '018' is_row-row lv_text lv_col space CHANGING ct_messages.
    CLEAR cv_value.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
FORM check_length USING    iv_row      TYPE i
                           iv_col      TYPE i
                           iv_text     TYPE csequence
                           iv_max      TYPE i
                  CHANGING ct_messages TYPE bapiret2_t.

  DATA lv_col TYPE string.

  IF iv_max > 0 AND strlen( iv_text ) > iv_max.
    PERFORM column_name USING iv_col CHANGING lv_col.
    PERFORM add_msg USING 'E' '022' iv_row lv_col iv_max space CHANGING ct_messages.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Split rows by packing list number (carry-forward)
*&---------------------------------------------------------------------*
FORM split_by_packing_list USING    it_rows     TYPE ty_row_tab
                           CHANGING ct_split    TYPE ty_pl_rows_tab
                                    ct_messages TYPE bapiret2_t.

  DATA lv_current TYPE zsd_packno.
  DATA lv_packno  TYPE string.
  DATA lv_col     TYPE string.

  CLEAR ct_split.

  LOOP AT it_rows INTO DATA(ls_row).

    PERFORM get_cell USING ls_row gs_col-packno CHANGING lv_packno.
    IF lv_packno IS NOT INITIAL.
      IF strlen( lv_packno ) > 10.
        PERFORM column_name USING gs_col-packno CHANGING lv_col.
        PERFORM add_msg USING 'E' '022' ls_row-row lv_col 10 space CHANGING ct_messages.
        CONTINUE.
      ENDIF.
      lv_current = |{ lv_packno ALPHA = IN WIDTH = 10 }|.
    ELSEIF lv_current IS INITIAL.
      " Data above the first packing list number cannot be assigned
      PERFORM add_msg USING 'E' '055' ls_row-row space space space CHANGING ct_messages.
      CONTINUE.
    ENDIF.

    READ TABLE ct_split ASSIGNING FIELD-SYMBOL(<ls_split>) WITH KEY packno = lv_current.
    IF sy-subrc <> 0.
      APPEND VALUE #( packno = lv_current ) TO ct_split ASSIGNING <ls_split>.
    ENDIF.
    APPEND ls_row TO <ls_split>-rows.

  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Header data, invoices and mandatory checks of one packing list
*&---------------------------------------------------------------------*
FORM map_header USING    it_rows TYPE ty_row_tab
                CHANGING cs_pl   TYPE ty_packing_list.

  TYPES: BEGIN OF ty_first,
           col   TYPE i,
           row   TYPE i,
           value TYPE string,
         END OF ty_first.
  DATA lt_first     TYPE SORTED TABLE OF ty_first WITH UNIQUE KEY col.
  DATA lv_value     TYPE string.
  DATA lv_inv       TYPE string.
  DATA lt_vbeln     TYPE ty_vbeln_tab.
  DATA lt_invalid   TYPE ty_text_tab.
  DATA lv_inv_found TYPE i.
  DATA lv_date      TYPE d.
  DATA lv_num       TYPE decfloat34.
  DATA lv_ok        TYPE abap_bool.
  DATA lv_col       TYPE string.
  DATA lv_gap       TYPE i.

  DATA(lt_map) = gt_map.
  DELETE lt_map WHERE col = 0.

  LOOP AT it_rows INTO DATA(ls_row).

    " Invoices: every row of the packing list may carry one or more
    PERFORM get_cell USING ls_row gs_col-invoice CHANGING lv_inv.
    IF lv_inv IS NOT INITIAL.
      PERFORM split_invoices USING lv_inv CHANGING lt_vbeln lt_invalid.
      lv_inv_found = lv_inv_found + lines( lt_vbeln ).
      INSERT LINES OF lt_vbeln INTO TABLE cs_pl-invoices.   " repeated = once
      LOOP AT lt_invalid INTO DATA(lv_invalid).
        PERFORM column_name USING gs_col-invoice CHANGING lv_col.
        PERFORM add_msg USING 'E' '018' ls_row-row lv_invalid lv_col space CHANGING cs_pl-messages.
      ENDLOOP.
    ENDIF.

    LOOP AT lt_map INTO DATA(ls_map).
      PERFORM get_cell USING ls_row ls_map-col CHANGING lv_value.
      IF lv_value IS INITIAL.
        CONTINUE.
      ENDIF.

      " The same value in several rows is fine, a different one is not
      READ TABLE lt_first INTO DATA(ls_first) WITH TABLE KEY col = ls_map-col.
      IF sy-subrc = 0.
        IF ls_first-value <> lv_value.
          PERFORM add_msg USING 'E' '020' ls_row-row ls_map-label ls_first-row space CHANGING cs_pl-messages.
        ENDIF.
        CONTINUE.
      ENDIF.
      INSERT VALUE #( col = ls_map-col row = ls_row-row value = lv_value ) INTO TABLE lt_first.

      ASSIGN COMPONENT ls_map-field OF STRUCTURE cs_pl-header TO FIELD-SYMBOL(<lv_field>).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      CASE ls_map-kind.
        WHEN 'D'.
          PERFORM parse_date USING lv_value CHANGING lv_date lv_ok.
          IF lv_ok = abap_true.
            <lv_field> = lv_date.
          ELSE.
            PERFORM column_name USING ls_map-col CHANGING lv_col.
            PERFORM add_msg USING 'E' '019' ls_row-row lv_value lv_col space CHANGING cs_pl-messages.
          ENDIF.
        WHEN 'A'.
          PERFORM parse_number USING lv_value CHANGING lv_num lv_ok.
          IF lv_ok = abap_true.
            <lv_field> = lv_num.
          ELSE.
            PERFORM column_name USING ls_map-col CHANGING lv_col.
            PERFORM add_msg USING 'E' '018' ls_row-row lv_value lv_col space CHANGING cs_pl-messages.
          ENDIF.
        WHEN OTHERS.
          " Multi-line cells (Alt+Enter) keep their line breaks
          PERFORM check_length USING ls_row-row ls_map-col lv_value ls_map-maxlen CHANGING cs_pl-messages.
          <lv_field> = lv_value.
      ENDCASE.
    ENDLOOP.

  ENDLOOP.

  DATA(lv_repeated) = lv_inv_found - lines( cs_pl-invoices ).
  IF lv_repeated > 0.
    PERFORM add_msg USING 'I' '056' lv_repeated space space space CHANGING cs_pl-messages.
  ENDIF.

  " Mandatory per packing list
  DATA(lv_first_row) = it_rows[ 1 ]-row.
  IF cs_pl-header-zsd_packdt IS INITIAL.
    PERFORM add_msg USING 'E' '017' lv_first_row 'Packing list date' space space CHANGING cs_pl-messages.
  ENDIF.
  IF cs_pl-invoices IS INITIAL.
    PERFORM add_msg USING 'E' '017' lv_first_row 'SAP Invoice Number' space space CHANGING cs_pl-messages.
  ENDIF.

  " Notify parties must be filled from 1 upwards (A10)
  DO gc_max_np TIMES.
    ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE cs_pl-header TO FIELD-SYMBOL(<lv_np>).
    IF <lv_np> IS INITIAL.
      IF lv_gap IS INITIAL.
        lv_gap = sy-index.
      ENDIF.
    ELSEIF lv_gap IS NOT INITIAL.
      PERFORM add_msg USING 'E' '031' sy-index lv_gap space space CHANGING cs_pl-messages.
    ENDIF.
  ENDDO.

ENDFORM.

*&---------------------------------------------------------------------*
*& Line items of one packing list
*&---------------------------------------------------------------------*
FORM map_lines USING    it_rows TYPE ty_row_tab
               CHANGING cs_pl   TYPE ty_packing_list.

  DATA lt_lines   TYPE ty_upload_line_tab.
  DATA ls_line    TYPE ty_upload_line.
  DATA lv_ignored TYPE i.
  DATA lv_text    TYPE string.
  DATA lv_num     TYPE decfloat34.
  DATA lv_dup     TYPE abap_bool.

  LOOP AT it_rows INTO DATA(ls_row).

    " A line item needs a Sl.No.; other rows (total row, header-only rows)
    " are not line items
    PERFORM get_cell USING ls_row gs_col-sn CHANGING lv_text.
    IF lv_text IS INITIAL.
      PERFORM get_cell USING ls_row gs_col-partno CHANGING lv_text.
      IF lv_text IS INITIAL.
        PERFORM get_cell USING ls_row gs_col-qty CHANGING lv_text.
      ENDIF.
      IF lv_text IS NOT INITIAL.
        lv_ignored = lv_ignored + 1.
      ENDIF.
      CONTINUE.
    ENDIF.

    CLEAR ls_line.
    ls_line-row = ls_row-row.

    PERFORM get_number USING ls_row gs_col-sn     abap_true  CHANGING lv_num cs_pl-messages.
    ls_line-sn     = lv_num.
    PERFORM get_number USING ls_row gs_col-inv_sn abap_true  CHANGING lv_num cs_pl-messages.
    ls_line-inv_sn = lv_num.
    PERFORM get_number USING ls_row gs_col-std    abap_true  CHANGING lv_num cs_pl-messages.
    ls_line-std    = lv_num.
    PERFORM get_number USING ls_row gs_col-art    abap_true  CHANGING lv_num cs_pl-messages.
    ls_line-art    = lv_num.
    PERFORM get_number USING ls_row gs_col-qty    abap_false CHANGING lv_num cs_pl-messages.
    ls_line-qty    = round( val = lv_num dec = 3 ).
    PERFORM get_number USING ls_row gs_col-gwt    abap_false CHANGING lv_num cs_pl-messages.
    ls_line-gwt    = round( val = lv_num dec = 3 ).
    PERFORM get_number USING ls_row gs_col-nwt    abap_false CHANGING lv_num cs_pl-messages.
    ls_line-nwt    = round( val = lv_num dec = 3 ).

    " Empty cell = merged cell, i.e. part of the group above (A06)
    PERFORM get_cell USING ls_row gs_col-art CHANGING lv_text.
    ls_line-art_filled = xsdbool( lv_text IS NOT INITIAL ).
    PERFORM get_cell USING ls_row gs_col-gwt CHANGING lv_text.
    ls_line-gwt_filled = xsdbool( lv_text IS NOT INITIAL ).
    PERFORM get_cell USING ls_row gs_col-nwt CHANGING lv_text.
    ls_line-nwt_filled = xsdbool( lv_text IS NOT INITIAL ).

    PERFORM get_cell USING ls_row gs_col-artno CHANGING lv_text.
    PERFORM check_length USING ls_row-row gs_col-artno lv_text 15 CHANGING cs_pl-messages.
    ls_line-artno = lv_text.

    PERFORM get_cell USING ls_row gs_col-partno CHANGING lv_text.
    PERFORM check_length USING ls_row-row gs_col-partno lv_text 40 CHANGING cs_pl-messages.
    ls_line-partno = to_upper( lv_text ).

    PERFORM get_cell USING ls_row gs_col-matdesc CHANGING lv_text.
    PERFORM check_length USING ls_row-row gs_col-matdesc lv_text 40 CHANGING cs_pl-messages.
    ls_line-matdesc = lv_text.

    PERFORM get_cell USING ls_row gs_col-hsn CHANGING lv_text.
    PERFORM check_length USING ls_row-row gs_col-hsn lv_text 17 CHANGING cs_pl-messages.
    ls_line-hsn = lv_text.

    PERFORM get_cell USING ls_row gs_col-pkg_type CHANGING lv_text.
    PERFORM check_length USING ls_row-row gs_col-pkg_type lv_text 20 CHANGING cs_pl-messages.
    ls_line-pkg_type = lv_text.

    IF ls_line-partno IS INITIAL.
      PERFORM add_msg USING 'E' '017' ls_row-row 'Part No.' space space CHANGING cs_pl-messages.
    ENDIF.
    PERFORM get_cell USING ls_row gs_col-qty CHANGING lv_text.
    IF lv_text IS INITIAL.
      PERFORM add_msg USING 'E' '017' ls_row-row 'Quantity' space space CHANGING cs_pl-messages.
    ENDIF.

    " Exact duplicate of an earlier row: counted once (A38). Same Sl.No.
    " with different content is an error raised in FORM build_items.
    lv_dup = abap_false.
    LOOP AT lt_lines INTO DATA(ls_prev) WHERE sn = ls_line-sn.
      DATA(ls_cmp) = ls_line.
      ls_cmp-row = ls_prev-row.
      IF ls_cmp = ls_prev.
        PERFORM add_msg USING 'W' '057' ls_row-row ls_prev-row space space CHANGING cs_pl-messages.
        lv_dup = abap_true.
      ENDIF.
      EXIT.
    ENDLOOP.
    IF lv_dup = abap_true.
      CONTINUE.
    ENDIF.

    APPEND ls_line TO lt_lines.
  ENDLOOP.

  IF lv_ignored > 0.
    PERFORM add_msg USING 'I' '050' lv_ignored space space space CHANGING cs_pl-messages.
  ENDIF.

  IF lt_lines IS INITIAL.
    PERFORM add_msg USING 'E' '036' space space space space CHANGING cs_pl-messages.
    RETURN.
  ENDIF.

  " Package values repeated on every row instead of merged cells (A06)
  PERFORM collapse_repeated_groups CHANGING lt_lines cs_pl-messages.
  PERFORM build_items USING cs_pl-packno lt_lines CHANGING cs_pl-items cs_pl-messages.

ENDFORM.

*&---------------------------------------------------------------------*
*& Step 2a: read all SAP tables ONCE for the invoices of all packing lists
*&---------------------------------------------------------------------*
FORM read_sap_data USING it_packing_lists TYPE ty_packing_list_tab.

  TYPES: BEGIN OF ty_area,
           kunnr TYPE kunnr,
           vkorg TYPE vkorg,
           vtweg TYPE vtweg,
           spart TYPE spart,
         END OF ty_area.
  DATA lt_inv    TYPE SORTED TABLE OF vbeln_vf WITH UNIQUE KEY table_line.
  DATA lt_area   TYPE SORTED TABLE OF ty_area WITH UNIQUE KEY kunnr vkorg vtweg spart.
  DATA lt_kunnr  TYPE SORTED TABLE OF kunnr WITH UNIQUE KEY table_line.
  DATA lt_bukrs  TYPE SORTED TABLE OF bukrs WITH UNIQUE KEY table_line.
  DATA lt_adrnr  TYPE ty_id_adrnr_tab.
  DATA lt_aubel  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.
  DATA lt_vgbel  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.
  DATA lt_zterm  TYPE SORTED TABLE OF dzterm WITH UNIQUE KEY table_line.
  DATA lt_matnr  TYPE SORTED TABLE OF matnr WITH UNIQUE KEY table_line.
  DATA lt_steuc  TYPE SORTED TABLE OF steuc WITH UNIQUE KEY table_line.
  DATA lt_branch TYPE SORTED TABLE OF ty_branch WITH UNIQUE KEY bukrs branch.
  DATA lv_matnr  TYPE matnr.
  DATA lv_steuc  TYPE steuc.

  " Invoices of all packing lists, each one once; Part No. -> material
  LOOP AT it_packing_lists INTO DATA(ls_pl).
    INSERT LINES OF ls_pl-invoices INTO TABLE lt_inv.
    LOOP AT ls_pl-items INTO DATA(ls_item) WHERE zsd_partno IS NOT INITIAL.
      READ TABLE gt_matnr TRANSPORTING NO FIELDS WITH TABLE KEY partno = ls_item-zsd_partno.
      IF sy-subrc <> 0.
        PERFORM matnr_of USING ls_item-zsd_partno CHANGING lv_matnr.
        INSERT VALUE #( partno = ls_item-zsd_partno matnr = lv_matnr ) INTO TABLE gt_matnr.
      ENDIF.
      IF ls_item-zsd_hsn IS NOT INITIAL.
        PERFORM steuc_of USING ls_item-zsd_hsn CHANGING lv_steuc.
        INSERT lv_steuc INTO TABLE lt_steuc.
      ENDIF.
    ENDLOOP.
  ENDLOOP.
  IF lt_inv IS INITIAL.
    RETURN.
  ENDIF.

  " Billing documents - logic B3, B6, B11, B20, B26, B43, B44
  SELECT vbeln, fkart, fksto, sfakn, bukrs, vkorg, vtweg, spart,
         kunag, zterm, inco1, inco2, bupla
    FROM vbrk
    FOR ALL ENTRIES IN @lt_inv
    WHERE vbeln = @lt_inv-table_line
    INTO TABLE @gt_vbrk.

  " Billing items - logic B13 (AUBEL), B42 (MATNR)
  SELECT vbeln, posnr, matnr, werks, aubel
    FROM vbrp
    FOR ALL ENTRIES IN @lt_inv
    WHERE vbeln = @lt_inv-table_line
    INTO TABLE @gt_vbrp.

  LOOP AT gt_vbrk INTO DATA(ls_vbrk).
    INSERT VALUE #( kunnr = ls_vbrk-kunag vkorg = ls_vbrk-vkorg
                    vtweg = ls_vbrk-vtweg spart = ls_vbrk-spart ) INTO TABLE lt_area.
    INSERT ls_vbrk-bukrs INTO TABLE lt_bukrs.
    INSERT VALUE #( bukrs = ls_vbrk-bukrs branch = ls_vbrk-bupla ) INTO TABLE lt_branch.
    INSERT ls_vbrk-zterm INTO TABLE lt_zterm.
  ENDLOOP.

  LOOP AT gt_vbrp INTO DATA(ls_vbrp).
    IF ls_vbrp-aubel IS NOT INITIAL.
      INSERT ls_vbrp-aubel INTO TABLE lt_aubel.
    ENDIF.
    IF ls_vbrp-matnr IS NOT INITIAL.
      INSERT ls_vbrp-matnr INTO TABLE lt_matnr.
    ENDIF.
  ENDLOOP.

  " Partners - logic B20 / B26: KNVP with KUNNR = KUNAG, PARVW SP / SH
  IF lt_area IS NOT INITIAL.
    SELECT kunnr, vkorg, vtweg, spart, parvw, parza, kunn2, defpa
      FROM knvp
      FOR ALL ENTRIES IN @lt_area
      WHERE kunnr = @lt_area-kunnr
        AND vkorg = @lt_area-vkorg
        AND vtweg = @lt_area-vtweg
        AND spart = @lt_area-spart
        AND ( parvw = @gc_parvw_sold OR parvw = @gc_parvw_ship )
      INTO TABLE @gt_knvp.
  ENDIF.

  LOOP AT gt_knvp INTO DATA(ls_knvp).
    INSERT ls_knvp-kunn2 INTO TABLE lt_kunnr.
  ENDLOOP.

  IF lt_kunnr IS NOT INITIAL.
    SELECT kunnr AS id, adrnr FROM kna1
      FOR ALL ENTRIES IN @lt_kunnr
      WHERE kunnr = @lt_kunnr-table_line
      INTO TABLE @gt_kna1.
  ENDIF.

  IF lt_bukrs IS NOT INITIAL.
    " Exporter address - logic B3 / B11
    SELECT bukrs AS id, adrnr FROM t001
      FOR ALL ENTRIES IN @lt_bukrs
      WHERE bukrs = @lt_bukrs-table_line
      INTO TABLE @gt_t001.

    " CIN - logic B7
    SELECT bukrs, paval FROM t001z
      FOR ALL ENTRIES IN @lt_bukrs
      WHERE bukrs = @lt_bukrs-table_line
        AND party = @gc_cin_party
      INTO TABLE @gt_t001z.
  ENDIF.

  " GST No. - logic B6
  IF lt_branch IS NOT INITIAL.
    SELECT bukrs, branch, gstin FROM j_1bbranch
      FOR ALL ENTRIES IN @lt_branch
      WHERE bukrs  = @lt_branch-bukrs
        AND branch = @lt_branch-branch
      INTO TABLE @gt_branch.
  ENDIF.

  LOOP AT gt_kna1 INTO DATA(ls_id).
    INSERT VALUE #( id = ls_id-adrnr adrnr = ls_id-adrnr ) INTO TABLE lt_adrnr.
  ENDLOOP.
  LOOP AT gt_t001 INTO ls_id.
    INSERT VALUE #( id = ls_id-adrnr adrnr = ls_id-adrnr ) INTO TABLE lt_adrnr.
  ENDLOOP.
  DELETE lt_adrnr WHERE adrnr IS INITIAL.
  PERFORM read_addresses USING lt_adrnr.

  " Order no. & date - logic B13: VBRP-AUBEL -> VBAK-VGBEL -> VBKD
  IF lt_aubel IS NOT INITIAL.
    SELECT vbeln, vgbel FROM vbak
      FOR ALL ENTRIES IN @lt_aubel
      WHERE vbeln = @lt_aubel-table_line
      INTO TABLE @gt_vbak.
    LOOP AT gt_vbak INTO DATA(ls_vbak) WHERE vgbel IS NOT INITIAL.
      INSERT ls_vbak-vgbel INTO TABLE lt_vgbel.
    ENDLOOP.
  ENDIF.
  IF lt_vgbel IS NOT INITIAL.
    " Header business data of the quotation (item 000000)
    SELECT vbeln, bstkd, bstdk FROM vbkd
      FOR ALL ENTRIES IN @lt_vgbel
      WHERE vbeln = @lt_vgbel-table_line
        AND posnr = '000000'
      INTO TABLE @gt_vbkd.
  ENDIF.

  " Payment term text - logic B43
  IF lt_zterm IS NOT INITIAL.
    SELECT zterm, ztagg, vtext FROM tvzbt
      FOR ALL ENTRIES IN @lt_zterm
      WHERE spras = @gc_langu
        AND zterm = @lt_zterm-table_line
      INTO TABLE @gt_tvzbt.
  ENDIF.

  " Materials: billed materials (B42) + Part No. of the items (B54, B58)
  LOOP AT gt_matnr INTO DATA(ls_matnr) WHERE matnr IS NOT INITIAL.
    INSERT ls_matnr-matnr INTO TABLE lt_matnr.
  ENDLOOP.
  IF lt_matnr IS NOT INITIAL.
    SELECT matnr, maktx FROM makt
      FOR ALL ENTRIES IN @lt_matnr
      WHERE matnr = @lt_matnr-table_line
        AND spras = @gc_langu
      INTO TABLE @gt_makt.
  ENDIF.

  " HSN of the billed material in the billing plant (B42, B58)
  IF gt_vbrp IS NOT INITIAL.
    SELECT matnr, werks, steuc FROM marc
      FOR ALL ENTRIES IN @gt_vbrp
      WHERE matnr = @gt_vbrp-matnr
        AND werks = @gt_vbrp-werks
      INTO TABLE @gt_marc.
  ENDIF.

  LOOP AT gt_marc INTO DATA(ls_marc) WHERE steuc IS NOT INITIAL.
    INSERT ls_marc-steuc INTO TABLE lt_steuc.
  ENDLOOP.

  " HSN texts - logic B42: T604N with SPRAS = EN, LAND1 = IN
  IF lt_steuc IS NOT INITIAL.
    SELECT steuc, text1 FROM t604n
      FOR ALL ENTRIES IN @lt_steuc
      WHERE spras = @gc_langu
        AND land1 = @gc_hsn_country
        AND steuc = @lt_steuc-table_line
      INTO TABLE @gt_t604n.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Addresses, e-mail, phones, country / region texts (once per run)
*&---------------------------------------------------------------------*
FORM read_addresses USING it_adrnr TYPE ty_id_adrnr_tab.

  TYPES: BEGIN OF ty_region,
           land1 TYPE land1,
           bland TYPE regio,
         END OF ty_region.
  DATA lt_adrc   TYPE STANDARD TABLE OF ty_adrc WITH EMPTY KEY.
  DATA lt_land   TYPE SORTED TABLE OF land1 WITH UNIQUE KEY table_line.
  DATA lt_region TYPE SORTED TABLE OF ty_region WITH UNIQUE KEY land1 bland.

  IF it_adrnr IS INITIAL.
    RETURN.
  ENDIF.

  " Logic B3 / B20 / B26 - address fields
  SELECT addrnumber, date_from, name1, street, str_suppl1, city1, city2,
         post_code1, region, country, building
    FROM adrc
    FOR ALL ENTRIES IN @it_adrnr
    WHERE addrnumber = @it_adrnr-adrnr
      AND nation     = @space
      AND date_from <= @sy-datum
    INTO TABLE @lt_adrc.

  " Latest valid version of each address
  SORT lt_adrc BY addrnumber date_from DESCENDING.
  DELETE ADJACENT DUPLICATES FROM lt_adrc COMPARING addrnumber.
  gt_adrc = lt_adrc.

  " Logic B21 / B27 - e-mail
  SELECT addrnumber, flgdefault, consnumber, smtp_addr
    FROM adr6
    FOR ALL ENTRIES IN @it_adrnr
    WHERE addrnumber = @it_adrnr-adrnr
      AND persnumber = @space
    INTO TABLE @gt_adr6.

  " Logic B22 / B23 / B28 / B29 - telephone (R3_USER 1), mobile (3)
  SELECT addrnumber, r3_user, flgdefault, consnumber, telnr_long
    FROM adr2
    FOR ALL ENTRIES IN @it_adrnr
    WHERE addrnumber = @it_adrnr-adrnr
      AND persnumber = @space
      AND ( r3_user = '1' OR r3_user = '3' )
    INTO TABLE @gt_adr2.

  " Logic B4 / B5 - country and region texts
  LOOP AT gt_adrc INTO DATA(ls_adrc) WHERE country IS NOT INITIAL.
    INSERT ls_adrc-country INTO TABLE lt_land.
    IF ls_adrc-region IS NOT INITIAL.
      INSERT VALUE #( land1 = ls_adrc-country bland = ls_adrc-region ) INTO TABLE lt_region.
    ENDIF.
  ENDLOOP.

  IF lt_land IS NOT INITIAL.
    SELECT land1, landx FROM t005t
      FOR ALL ENTRIES IN @lt_land
      WHERE spras = @gc_langu
        AND land1 = @lt_land-table_line
      INTO TABLE @gt_t005t.
  ENDIF.

  IF lt_region IS NOT INITIAL.
    SELECT land1, bland, bezei FROM t005u
      FOR ALL ENTRIES IN @lt_region
      WHERE spras = @gc_langu
        AND land1 = @lt_region-land1
        AND bland = @lt_region-bland
      INTO TABLE @gt_t005u.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Step 2b: form data of every packing list (internal tables only)
*&---------------------------------------------------------------------*
FORM prepare_all.

  DATA ls_result TYPE ty_result.

  CLEAR gt_result.
  LOOP AT gt_packing_lists INTO DATA(ls_pl).
    PERFORM build_one USING ls_pl CHANGING ls_result.
    INSERT ls_result INTO TABLE gt_result.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Build the form data of ONE packing list from its own invoices and items
*&---------------------------------------------------------------------*
FORM build_one USING    is_pl     TYPE ty_packing_list
               CHANGING cs_result TYPE ty_result.

  DATA lt_msg        TYPE bapiret2_t.
  DATA lt_vbrk       TYPE ty_vbrk_list.
  DATA lt_vbrp       TYPE ty_vbrp_tab.
  DATA lt_items      TYPE ty_item_tab.
  DATA lt_headings   TYPE ty_hsn_heading_tab.
  DATA lt_goods      TYPE ty_text_tab.
  DATA lt_lines      TYPE ty_text_tab.
  DATA ls_vbrk       TYPE ty_vbrk.
  DATA ls_sold_to    TYPE ty_partner.
  DATA ls_ship_to    TYPE ty_partner.
  DATA ls_party      TYPE ty_party_data.
  DATA ls_header     TYPE zsd_s_epack_af_hdr.
  DATA ls_totals     TYPE ty_totals.
  DATA lv_np_count   TYPE i.
  DATA lv_errors     TYPE abap_bool.
  DATA lv_text       TYPE string.
  DATA lv_vtext      TYPE tvzbt-vtext.
  DATA lv_buyer_ref  TYPE string.

  CLEAR cs_result.
  lt_msg = is_pl-messages.

  cs_result-packno     = is_pl-packno.
  cs_result-inv_count  = lines( is_pl-invoices ).
  cs_result-item_count = lines( is_pl-items ).
  cs_result-print-form = gc_form_name.

  " Errors from the upload stop this packing list - but no other one
  PERFORM has_errors USING lt_msg CHANGING lv_errors.
  IF lv_errors = abap_false.

    PERFORM check_invoices    USING is_pl   CHANGING lt_vbrk lt_msg.
    PERFORM check_consistency USING lt_vbrk CHANGING lt_msg.
    PERFORM check_authority   USING lt_vbrk CHANGING lt_msg.

    " FS 2.1 - GW >= NW at output time
    READ TABLE lt_msg TRANSPORTING NO FIELDS WITH KEY number = '001'.
    IF sy-subrc <> 0.
      PERFORM check_weights USING is_pl-items CHANGING lt_msg.
    ENDIF.

    PERFORM count_notify_parties USING is_pl-header CHANGING lv_np_count.

    " All invoices agree (check_consistency), so the header values of
    " any of them are the values of all of them
    PERFORM has_errors USING lt_msg CHANGING lv_errors.
    IF lv_errors = abap_false AND lt_vbrk IS NOT INITIAL.
      READ TABLE lt_vbrk INTO ls_vbrk INDEX 1.
      " Logic B20 / B26 - a missing partner is an error (no consignee)
      PERFORM get_partner USING ls_vbrk gc_parvw_sold CHANGING ls_sold_to lt_msg.
      PERFORM get_partner USING ls_vbrk gc_parvw_ship CHANGING ls_ship_to lt_msg.
    ENDIF.
  ENDIF.

  PERFORM has_errors USING lt_msg CHANGING lv_errors.
  IF lv_errors = abap_true.
    cs_result-status   = gc_st_error.
    cs_result-messages = lt_msg.
    RETURN.
  ENDIF.

  " Billing items of THIS packing list's invoices only
  LOOP AT is_pl-invoices INTO DATA(lv_vbeln).
    LOOP AT gt_vbrp INTO DATA(ls_vbrp) WHERE vbeln = lv_vbeln.
      INSERT ls_vbrp INTO TABLE lt_vbrp.
    ENDLOOP.
  ENDLOOP.

  READ TABLE gt_t001 INTO DATA(ls_t001) WITH TABLE KEY id = ls_vbrk-bukrs.
  READ TABLE gt_adrc INTO DATA(ls_company) WITH TABLE KEY addrnumber = ls_t001-adrnr.
  IF sy-subrc <> 0.
    CLEAR ls_company.
  ENDIF.
  READ TABLE gt_branch INTO DATA(ls_branch) WITH TABLE KEY bukrs = ls_vbrk-bukrs branch = ls_vbrk-bupla.
  IF sy-subrc <> 0.
    CLEAR ls_branch.
  ENDIF.
  READ TABLE gt_t001z INTO DATA(ls_t001z) WITH TABLE KEY bukrs = ls_vbrk-bukrs.
  IF sy-subrc <> 0.
    CLEAR ls_t001z.
  ENDIF.
  READ TABLE gt_adrc INTO DATA(ls_ship_adr) WITH TABLE KEY addrnumber = ls_ship_to-adrnr.
  IF sy-subrc <> 0.
    CLEAR ls_ship_adr.
  ENDIF.

  " ---- Document ---------------------------------------------------------
  ls_header-packno = is_pl-header-zsd_packno.                 " logic B9
  PERFORM format_date USING is_pl-header-zsd_packdt CHANGING lv_text.
  ls_header-packdt   = lv_text.                               " logic B10
  ls_header-np_count = lv_np_count.
  ls_header-format   = COND #( WHEN lv_np_count > 0 THEN gc_fmt_with_np ELSE gc_fmt_wout_np ).

  " ---- Exporter - logic B3-B7, B11 --------------------------------------
  PERFORM exporter_data USING ls_t001-adrnr CHANGING ls_party lt_msg.
  ls_header-exp_name    = ls_party-name.
  ls_header-exp_address = ls_party-address.
  ls_header-sign_for    = ls_party-name.
  ls_header-exp_gstin   = ls_branch-gstin.                    " logic B6
  ls_header-exp_cin     = ls_t001z-paval.                     " logic B7
  " Logic B11: company address field BUILDING; the form prints "IEC:"
  lv_text = ls_company-building.
  REPLACE FIRST OCCURRENCE OF REGEX '^\s*IEC\s*[:.\-]?\s*' IN lv_text WITH `` IGNORING CASE.
  ls_header-exp_iec = lv_text.

  IF ls_header-exp_gstin IS INITIAL.
    PERFORM add_msg USING 'W' '049' 'GST No.' space space space CHANGING lt_msg.
  ENDIF.
  IF ls_header-exp_cin IS INITIAL.
    PERFORM add_msg USING 'W' '049' 'CIN No.' space space space CHANGING lt_msg.
  ENDIF.
  IF ls_header-exp_iec IS INITIAL.
    PERFORM add_msg USING 'W' '049' 'IEC (Exporter''s Ref.)' space space space CHANGING lt_msg.
  ENDIF.

  " ---- Consignee (ship-to, B26-B29) and buyer (sold-to, B20-B23) ---------
  " Contact person and FAX are not in the logic sheet - left empty (Q09)
  PERFORM party_data USING ls_ship_to CHANGING ls_party lt_msg.
  ls_header-cns_name    = ls_party-name.
  ls_header-cns_address = ls_party-address.
  ls_header-cns_tel     = ls_party-tel.
  ls_header-cns_mobile  = ls_party-mobile.
  ls_header-cns_email   = ls_party-email.

  PERFORM party_data USING ls_sold_to CHANGING ls_party lt_msg.
  ls_header-buy_name    = ls_party-name.
  ls_header-buy_address = ls_party-address.
  ls_header-buy_tel     = ls_party-tel.
  ls_header-buy_mobile  = ls_party-mobile.
  ls_header-buy_email   = ls_party-email.
  ls_header-same_party  = xsdbool( ls_sold_to-kunnr = ls_ship_to-kunnr ).

  " ---- Notify parties 1-5 - logic B31 ------------------------------------
  PERFORM text_block USING is_pl-header-zsd_np1 CHANGING ls_header-np1.
  PERFORM text_block USING is_pl-header-zsd_np2 CHANGING ls_header-np2.
  PERFORM text_block USING is_pl-header-zsd_np3 CHANGING ls_header-np3.
  PERFORM text_block USING is_pl-header-zsd_np4 CHANGING ls_header-np4.
  PERFORM text_block USING is_pl-header-zsd_np5 CHANGING ls_header-np5.

  " ---- References and countries - logic B13, B14, B16, B17 --------------
  PERFORM order_reference USING lt_vbrp CHANGING lt_lines lt_msg.
  ls_header-order_ref = concat_lines_of( table = lt_lines sep = cl_abap_char_utilities=>newline ).

  lv_buyer_ref = condense( CONV string( is_pl-header-zsd_party_ref ) ).
  IF is_pl-header-zsd_ref_dt IS NOT INITIAL.
    PERFORM format_date USING is_pl-header-zsd_ref_dt CHANGING lv_text.
    lv_buyer_ref = condense( |{ lv_buyer_ref } DATE: { lv_text }| ).
  ENDIF.
  ls_header-buyer_ref = lv_buyer_ref.

  ls_header-cnty_orgn = is_pl-header-zsd_cnty_orgn.
  PERFORM country_text USING ls_ship_adr-country CHANGING lv_text.
  ls_header-cnty_dest = lv_text.

  " ---- Shipping - logic B33-B38 ------------------------------------------
  ls_header-pre_carr  = is_pl-header-zsd_prec.
  ls_header-rec_place = is_pl-header-zsd_rec_pc.
  ls_header-vessel    = is_pl-header-zsd_vsl_flt.
  ls_header-pol       = is_pl-header-zsd_pol.
  ls_header-pod       = is_pl-header-zsd_pod.
  ls_header-pld       = is_pl-header-zsd_pld.

  " ---- Terms - logic B43 (payment term text, lowest ZTAGG), B44 ----------
  LOOP AT gt_tvzbt INTO DATA(ls_tvzbt) WHERE zterm = ls_vbrk-zterm.
    lv_vtext = ls_tvzbt-vtext.
    EXIT.
  ENDLOOP.
  ls_header-pay_terms  = |PAYMENT TERM: { COND string( WHEN lv_vtext IS NOT INITIAL THEN lv_vtext ELSE ls_vbrk-zterm ) }|.
  ls_header-ship_terms = condense( |{ ls_vbrk-inco1 } { ls_vbrk-inco2 }| ).

  " ---- Description of goods (B42) and line items (B47-B58) --------------
  PERFORM goods_lines   USING is_pl lt_vbrp CHANGING lt_goods.
  PERFORM resolve_items USING is_pl lt_vbrp lt_goods CHANGING lt_items lt_headings lt_msg.
  ls_header-goods_desc = concat_lines_of( table = lt_goods sep = cl_abap_char_utilities=>newline ).

  " ---- Marks, containers, packages - logic B39-B41 -----------------------
  PERFORM calc_totals USING lt_items CHANGING ls_totals.
  PERFORM text_block USING is_pl-header-zsd_mark CHANGING ls_header-marks.
  PERFORM text_block USING is_pl-header-zsd_cont CHANGING ls_header-containers.
  PERFORM text_block USING is_pl-header-zsd_nopack CHANGING ls_header-no_pkgs.
  IF ls_header-no_pkgs IS INITIAL.
    PERFORM packages_text USING ls_totals CHANGING ls_header-no_pkgs.
  ENDIF.

  " ---- Totals row and declaration (B52, B61) -----------------------------
  ls_header-tot_art = |{ ls_totals-articles }|.
  PERFORM format_quantity USING ls_totals-quantity CHANGING lv_text.
  ls_header-tot_qty = lv_text.
  PERFORM format_weight USING ls_totals-gross CHANGING lv_text.
  ls_header-tot_gwt = lv_text.
  PERFORM format_weight USING ls_totals-net CHANGING lv_text.
  ls_header-tot_nwt = lv_text.
  PERFORM total_text USING ls_totals CHANGING lv_text.
  ls_header-tot_text = |({ lv_text })|.
  PERFORM text_block USING is_pl-header-zsd_pl_decl CHANGING ls_header-declaration.

  cs_result-bukrs        = ls_vbrk-bukrs.
  cs_result-print-header = ls_header.
  PERFORM build_print_items USING lt_items lt_headings CHANGING cs_result-print-items.

  cs_result-messages = lt_msg.
  READ TABLE lt_msg TRANSPORTING NO FIELDS WITH KEY type = 'W'.
  cs_result-status = COND #( WHEN sy-subrc = 0 THEN gc_st_warning ELSE gc_st_success ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Invoices exist and are not cancelled
*&---------------------------------------------------------------------*
FORM check_invoices USING    is_pl       TYPE ty_packing_list
                    CHANGING ct_vbrk     TYPE ty_vbrk_list
                             ct_messages TYPE bapiret2_t.

  DATA lv_errors TYPE abap_bool.

  CLEAR ct_vbrk.

  LOOP AT is_pl-invoices INTO DATA(lv_vbeln).
    DATA(lv_ext) = |{ lv_vbeln ALPHA = OUT }|.
    READ TABLE gt_vbrk INTO DATA(ls_vbrk) WITH TABLE KEY vbeln = lv_vbeln.
    IF sy-subrc <> 0.
      PERFORM add_msg USING 'E' '008' lv_ext space space space CHANGING ct_messages.
      CONTINUE.
    ENDIF.
    IF ls_vbrk-fksto = abap_true OR ls_vbrk-sfakn IS NOT INITIAL.
      PERFORM add_msg USING 'E' '009' lv_ext space space space CHANGING ct_messages.
      CONTINUE.
    ENDIF.
    INSERT ls_vbrk INTO TABLE ct_vbrk.
  ENDLOOP.

  PERFORM has_errors USING ct_messages CHANGING lv_errors.
  IF ct_vbrk IS INITIAL AND lv_errors = abap_false.
    PERFORM add_msg USING 'E' '047' is_pl-packno space space space CHANGING ct_messages.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& One packing list prints ONE exporter, GSTIN, buyer, consignee, payment
*& term and Incoterm. Invoices that disagree are rejected rather than
*& printing the value of an arbitrary invoice (A04).
*&---------------------------------------------------------------------*
FORM check_consistency USING    it_vbrk     TYPE ty_vbrk_list
                       CHANGING ct_messages TYPE bapiret2_t.

  DATA ls_first TYPE ty_vbrk.
  DATA lv_inv   TYPE string.
  DATA lv_diff  TYPE string.

  IF lines( it_vbrk ) < 2.
    RETURN.
  ENDIF.

  READ TABLE it_vbrk INTO ls_first INDEX 1.

  LOOP AT it_vbrk INTO DATA(ls_vbrk) FROM 2.
    lv_inv = |{ ls_first-vbeln ALPHA = OUT }/{ ls_vbrk-vbeln ALPHA = OUT }|.
    IF ls_vbrk-bukrs <> ls_first-bukrs.
      lv_diff = |{ ls_first-bukrs } / { ls_vbrk-bukrs }|.
      PERFORM add_msg USING 'E' '011' 'Company code' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
    IF ls_vbrk-bupla <> ls_first-bupla.
      lv_diff = |{ ls_first-bupla } / { ls_vbrk-bupla }|.
      PERFORM add_msg USING 'E' '011' 'Business place (GSTIN)' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
    IF ls_vbrk-kunag <> ls_first-kunag.
      lv_diff = |{ ls_first-kunag ALPHA = OUT } / { ls_vbrk-kunag ALPHA = OUT }|.
      PERFORM add_msg USING 'E' '011' 'Sold-to party' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
    IF ls_vbrk-vkorg <> ls_first-vkorg OR ls_vbrk-vtweg <> ls_first-vtweg OR ls_vbrk-spart <> ls_first-spart.
      lv_diff = |{ ls_first-vkorg }/{ ls_first-vtweg }/{ ls_first-spart } / { ls_vbrk-vkorg }/{ ls_vbrk-vtweg }/{ ls_vbrk-spart }|.
      PERFORM add_msg USING 'E' '011' 'Sales area' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
    IF ls_vbrk-zterm <> ls_first-zterm.
      lv_diff = |{ ls_first-zterm } / { ls_vbrk-zterm }|.
      PERFORM add_msg USING 'E' '011' 'Payment terms' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
    IF ls_vbrk-inco1 <> ls_first-inco1 OR ls_vbrk-inco2 <> ls_first-inco2.
      lv_diff = |{ ls_first-inco1 } { ls_first-inco2 } / { ls_vbrk-inco1 } { ls_vbrk-inco2 }|.
      PERFORM add_msg USING 'E' '011' 'Incoterms' lv_diff lv_inv space CHANGING ct_messages.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Print authorization per sales org. of the invoices (V_VBRK_VKO), buffered
*&---------------------------------------------------------------------*
FORM check_authority USING    it_vbrk     TYPE ty_vbrk_list
                     CHANGING ct_messages TYPE bapiret2_t.

  DATA ls_auth TYPE ty_auth.
  DATA lt_msg  TYPE bapiret2_t.

  LOOP AT it_vbrk INTO DATA(ls_vbrk).

    READ TABLE gt_auth INTO ls_auth WITH TABLE KEY object = 'V_VBRK_VKO' value = ls_vbrk-vkorg.
    IF sy-subrc <> 0.
      AUTHORITY-CHECK OBJECT 'V_VBRK_VKO'
        ID 'VKORG' FIELD ls_vbrk-vkorg
        ID 'ACTVT' FIELD gc_actvt_print.
      ls_auth = VALUE #( object = 'V_VBRK_VKO' value = ls_vbrk-vkorg ok = xsdbool( sy-subrc = 0 ) ).
      INSERT ls_auth INTO TABLE gt_auth.
    ENDIF.
    IF ls_auth-ok = abap_false.
      PERFORM add_msg USING 'E' '012' 'Sales org.' ls_vbrk-vkorg gc_actvt_print space CHANGING lt_msg.
    ENDIF.

  ENDLOOP.

  " Each missing authorization once per packing list
  SORT lt_msg BY message_v1 message_v2.
  DELETE ADJACENT DUPLICATES FROM lt_msg COMPARING message_v1 message_v2.
  APPEND LINES OF lt_msg TO ct_messages.

ENDFORM.

*&---------------------------------------------------------------------*
*& Logic B20 / B26: KNVP with KUNNR = KUNAG and the sales area of the
*& invoice, PARVW = SP / SH -> KUNN2. Several partners: default (DEFPA),
*& else the first by counter PARZA (A31).
*&---------------------------------------------------------------------*
FORM get_partner USING    is_vbrk     TYPE ty_vbrk
                          iv_parvw    TYPE parvw
                 CHANGING cs_partner  TYPE ty_partner
                          ct_messages TYPE bapiret2_t.

  DATA lv_count TYPE i.
  DATA ls_hit   TYPE ty_knvp.
  DATA lv_cust  TYPE string.
  DATA lv_area  TYPE string.
  DATA lv_used  TYPE string.

  CLEAR cs_partner.

  LOOP AT gt_knvp INTO DATA(ls_knvp)
       WHERE kunnr = is_vbrk-kunag
         AND vkorg = is_vbrk-vkorg
         AND vtweg = is_vbrk-vtweg
         AND spart = is_vbrk-spart
         AND parvw = iv_parvw.
    lv_count = lv_count + 1.
    IF ls_hit IS INITIAL OR ( ls_knvp-defpa = abap_true AND ls_hit-defpa = abap_false ).
      ls_hit = ls_knvp.
    ENDIF.
  ENDLOOP.

  DATA(lv_function) = COND string( WHEN iv_parvw = gc_parvw_sold THEN `Sold-to (SP)` ELSE `Ship-to (SH)` ).
  IF lv_count = 0.
    lv_cust = |{ is_vbrk-kunag ALPHA = OUT }|.
    lv_area = |{ is_vbrk-vkorg }/{ is_vbrk-vtweg }/{ is_vbrk-spart }|.
    PERFORM add_msg USING 'E' '042' lv_function lv_cust lv_area space CHANGING ct_messages.
    RETURN.
  ELSEIF lv_count > 1.
    lv_cust = |{ is_vbrk-kunag ALPHA = OUT }|.
    lv_used = |{ ls_hit-kunn2 ALPHA = OUT }|.
    PERFORM add_msg USING 'W' '060' lv_cust lv_count lv_function lv_used CHANGING ct_messages.
  ENDIF.

  cs_partner-kunnr = ls_hit-kunn2.
  READ TABLE gt_kna1 INTO DATA(ls_kna1) WITH TABLE KEY id = ls_hit-kunn2.
  IF sy-subrc = 0.
    cs_partner-adrnr = ls_kna1-adrnr.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& FS 1.1: one Adobe Form; it shows the notify-party section when notify
*& parties are maintained (A05)
*&---------------------------------------------------------------------*
FORM count_notify_parties USING    is_header   TYPE ty_header
                          CHANGING cv_np_count TYPE i.

  cv_np_count = 0.
  DO gc_max_np TIMES.
    ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE is_header TO FIELD-SYMBOL(<lv_np>).
    IF sy-subrc = 0 AND <lv_np> IS NOT INITIAL.
      cv_np_count = cv_np_count + 1.
    ENDIF.
  ENDDO.

ENDFORM.

*&---------------------------------------------------------------------*
*& Description of goods: uploaded text (ZSD_HSN_DESC) if filled (A14),
*& else logic B42: VBRP-MATNR -> MARC-STEUC -> T604N-TEXT1, "club common
*& HSN code and description in all billing documents"
*&---------------------------------------------------------------------*
FORM goods_lines USING    is_pl    TYPE ty_packing_list
                          it_vbrp  TYPE ty_vbrp_tab
                 CHANGING ct_lines TYPE ty_text_tab.

  DATA lt_seen TYPE SORTED TABLE OF steuc WITH UNIQUE KEY table_line.
  DATA lv_hsn  TYPE zsd_hsn.

  PERFORM split_lines USING is_pl-header-zsd_hsn_desc CHANGING ct_lines.
  IF ct_lines IS NOT INITIAL.
    RETURN.
  ENDIF.

  LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE matnr IS NOT INITIAL.
    READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = ls_vbrp-matnr werks = ls_vbrp-werks.
    IF sy-subrc <> 0 OR ls_marc-steuc IS INITIAL.
      CONTINUE.
    ENDIF.
    INSERT ls_marc-steuc INTO TABLE lt_seen.
    IF sy-subrc <> 0.
      CONTINUE.                                   " HSN already clubbed
    ENDIF.
    PERFORM format_hsn USING ls_marc-steuc CHANGING lv_hsn.
    READ TABLE gt_t604n INTO DATA(ls_t604n) WITH TABLE KEY steuc = ls_marc-steuc.
    IF sy-subrc = 0 AND ls_t604n-text1 IS NOT INITIAL.
      APPEND |{ ls_t604n-text1 } - HS CODE : { lv_hsn }| TO ct_lines.
    ELSE.
      APPEND |HS CODE : { lv_hsn }| TO ct_lines.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Line items: product description (B54) and HSN (B58) from the upload
*& OR the material master of the Part No. (step 53); section headings
*&---------------------------------------------------------------------*
FORM resolve_items USING    is_pl       TYPE ty_packing_list
                            it_vbrp     TYPE ty_vbrp_tab
                            it_goods    TYPE ty_text_tab
                   CHANGING ct_items    TYPE ty_item_tab
                            ct_headings TYPE ty_hsn_heading_tab
                            ct_messages TYPE bapiret2_t.

  DATA lv_sn    TYPE i.
  DATA lv_matnr TYPE matnr.
  DATA lv_werks TYPE werks_d.
  DATA lv_hsn   TYPE zsd_hsn.
  DATA lv_steuc TYPE steuc.
  DATA lv_text  TYPE string.
  DATA lv_hsntx TYPE string.

  ct_items = is_pl-items.
  CLEAR ct_headings.

  LOOP AT ct_items ASSIGNING FIELD-SYMBOL(<ls_item>).

    lv_sn = <ls_item>-zsd_sn.
    CLEAR: lv_matnr, lv_werks.
    READ TABLE gt_matnr INTO DATA(ls_matnr) WITH TABLE KEY partno = <ls_item>-zsd_partno.
    IF sy-subrc = 0.
      lv_matnr = ls_matnr-matnr.
    ENDIF.

    " Plant of the material in this packing list's invoices
    LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE matnr = lv_matnr.
      lv_werks = ls_vbrp-werks.
      EXIT.
    ENDLOOP.
    IF lv_werks IS INITIAL.
      " Reconciliation hint only - not a logic-sheet rule
      PERFORM add_msg USING 'W' '030' lv_sn <ls_item>-zsd_partno space space CHANGING ct_messages.
    ENDIF.

    " Logic B54: ZSD_MATDESC OR material text
    IF <ls_item>-zsd_matdesc IS INITIAL.
      READ TABLE gt_makt INTO DATA(ls_makt) WITH TABLE KEY matnr = lv_matnr.
      IF sy-subrc = 0.
        <ls_item>-zsd_matdesc = ls_makt-maktx.
      ELSE.
        PERFORM add_msg USING 'W' '048' lv_sn space space space CHANGING ct_messages.
      ENDIF.
    ENDIF.

    " Logic B58: ZSD_HSN OR MARC-STEUC of the Part No.
    IF <ls_item>-zsd_hsn IS INITIAL.
      READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = lv_matnr werks = lv_werks.
      IF sy-subrc = 0 AND ls_marc-steuc IS NOT INITIAL.
        <ls_item>-zsd_hsn = ls_marc-steuc.
      ELSE.
        PERFORM add_msg USING 'W' '044' lv_sn space space space CHANGING ct_messages.
      ENDIF.
    ENDIF.
    PERFORM format_hsn USING <ls_item>-zsd_hsn CHANGING lv_hsn.
    <ls_item>-zsd_hsn = lv_hsn.

  ENDLOOP.

  " Section heading per HSN, in order of first appearance
  LOOP AT ct_items INTO DATA(ls_item).
    READ TABLE ct_headings TRANSPORTING NO FIELDS WITH KEY hsn = ls_item-zsd_hsn.
    IF sy-subrc = 0.
      CONTINUE.
    ENDIF.
    PERFORM steuc_of USING ls_item-zsd_hsn CHANGING lv_steuc.
    CLEAR lv_hsntx.
    READ TABLE gt_t604n INTO DATA(ls_t604n) WITH TABLE KEY steuc = lv_steuc.
    IF sy-subrc = 0.
      lv_hsntx = ls_t604n-text1.
    ENDIF.
    PERFORM heading_for_hsn USING ls_item-zsd_hsn it_goods lv_hsntx CHANGING lv_text.
    APPEND VALUE #( hsn = ls_item-zsd_hsn text = lv_text ) TO ct_headings.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Logic B13: VBRP-AUBEL -> VBAK-VGBEL (quotation) -> VBKD-BSTKD / BSTDK.
*& Several orders: all distinct references, joined with " & " (A17)
*&---------------------------------------------------------------------*
FORM order_reference USING    it_vbrp     TYPE ty_vbrp_tab
                     CHANGING ct_lines    TYPE ty_text_tab
                              ct_messages TYPE bapiret2_t.

  DATA lt_bstkd TYPE ty_text_tab.
  DATA lt_bstdk TYPE ty_text_tab.
  DATA lt_seen  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.
  DATA lv_refs  TYPE string.
  DATA lv_dates TYPE string.
  DATA lv_val   TYPE string.

  CLEAR ct_lines.

  LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE aubel IS NOT INITIAL.
    INSERT ls_vbrp-aubel INTO TABLE lt_seen.
    IF sy-subrc <> 0.
      CONTINUE.
    ENDIF.

    READ TABLE gt_vbak INTO DATA(ls_vbak) WITH TABLE KEY vbeln = ls_vbrp-aubel.
    IF sy-subrc <> 0 OR ls_vbak-vgbel IS INITIAL.
      lv_val = |{ ls_vbrp-aubel ALPHA = OUT }|.
      PERFORM add_msg USING 'W' '061' lv_val space space space CHANGING ct_messages.
      CONTINUE.
    ENDIF.

    READ TABLE gt_vbkd INTO DATA(ls_vbkd) WITH TABLE KEY vbeln = ls_vbak-vgbel.
    IF sy-subrc <> 0.
      CONTINUE.
    ENDIF.

    lv_val = condense( CONV string( ls_vbkd-bstkd ) ).
    IF lv_val IS NOT INITIAL.
      READ TABLE lt_bstkd TRANSPORTING NO FIELDS WITH KEY table_line = lv_val.
      IF sy-subrc <> 0.
        APPEND lv_val TO lt_bstkd.
      ENDIF.
    ENDIF.
    IF ls_vbkd-bstdk IS NOT INITIAL.
      lv_val = |{ ls_vbkd-bstdk+6(2) }.{ ls_vbkd-bstdk+4(2) }.{ ls_vbkd-bstdk(4) }|.
      READ TABLE lt_bstdk TRANSPORTING NO FIELDS WITH KEY table_line = lv_val.
      IF sy-subrc <> 0.
        APPEND lv_val TO lt_bstdk.
      ENDIF.
    ENDIF.
  ENDLOOP.

  IF lt_bstkd IS INITIAL AND lt_bstdk IS INITIAL.
    RETURN.
  ENDIF.

  PERFORM join_parts USING lt_bstkd ` & ` CHANGING lv_refs.
  PERFORM join_parts USING lt_bstdk ` & ` CHANGING lv_dates.
  IF lv_dates IS NOT INITIAL.
    APPEND |{ lv_refs } DATE: { lv_dates }| TO ct_lines.
  ELSE.
    APPEND lv_refs TO ct_lines.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Logic B3-B5: NAME1; address = STREET, STR_SUPPL1, CITY1 - POST_CODE1,
*& REGION text, COUNTRY text (GST No. B6 is a separate header field)
*&---------------------------------------------------------------------*
FORM exporter_data USING    iv_adrnr    TYPE adrnr
                   CHANGING cs_party    TYPE ty_party_data
                            ct_messages TYPE bapiret2_t.

  DATA lt_lines   TYPE ty_text_tab.
  DATA lt_parts   TYPE ty_text_tab.
  DATA lv_text    TYPE string.
  DATA lv_country TYPE string.

  CLEAR cs_party.

  READ TABLE gt_adrc INTO DATA(ls_adr) WITH TABLE KEY addrnumber = iv_adrnr.
  IF sy-subrc <> 0.
    PERFORM add_msg USING 'W' '049' 'Exporter address' space space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  cs_party-name = ls_adr-name1.

  APPEND CONV string( ls_adr-street )     TO lt_lines.
  APPEND CONV string( ls_adr-str_suppl1 ) TO lt_lines.

  READ TABLE gt_t005u INTO DATA(ls_t005u) WITH TABLE KEY land1 = ls_adr-country bland = ls_adr-region.
  IF sy-subrc <> 0.
    CLEAR ls_t005u.
  ENDIF.
  PERFORM country_text USING ls_adr-country CHANGING lv_country.
  lt_parts = VALUE #( ( condense( |{ ls_adr-city1 }{ COND string( WHEN ls_adr-post_code1 IS NOT INITIAL
                                                                   THEN | - { ls_adr-post_code1 }| ) }| ) )
                      ( to_upper( ls_t005u-bezei ) )
                      ( lv_country ) ).
  PERFORM join_parts USING lt_parts `, ` CHANGING lv_text.
  APPEND lv_text TO lt_lines.

  DELETE lt_lines WHERE table_line IS INITIAL.
  cs_party-address = concat_lines_of( table = lt_lines sep = cl_abap_char_utilities=>newline ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Logic B20-B23 / B26-B29: NAME1, STREET, STR_SUPPL1, CITY1, CITY2,
*& POST_CODE1, COUNTRY text; e-mail; telephone (R3_USER 1) / mobile (3)
*&---------------------------------------------------------------------*
FORM party_data USING    is_partner  TYPE ty_partner
                CHANGING cs_party    TYPE ty_party_data
                         ct_messages TYPE bapiret2_t.

  DATA lt_parts   TYPE ty_text_tab.
  DATA lv_text    TYPE string.
  DATA lv_country TYPE string.

  CLEAR cs_party.

  READ TABLE gt_adrc INTO DATA(ls_adr) WITH TABLE KEY addrnumber = is_partner-adrnr.
  IF sy-subrc <> 0.
    IF is_partner-kunnr IS NOT INITIAL.
      lv_text = |Address of customer { is_partner-kunnr ALPHA = OUT }|.
      PERFORM add_msg USING 'W' '049' lv_text space space space CHANGING ct_messages.
    ENDIF.
    RETURN.
  ENDIF.

  cs_party-name = ls_adr-name1.

  PERFORM country_text USING ls_adr-country CHANGING lv_country.
  lt_parts = VALUE #( ( CONV #( ls_adr-street ) )
                      ( CONV #( ls_adr-str_suppl1 ) )
                      ( CONV #( ls_adr-city2 ) )
                      ( condense( |{ ls_adr-post_code1 } { ls_adr-city1 }| ) )
                      ( lv_country ) ).
  PERFORM join_parts USING lt_parts `, ` CHANGING cs_party-address.

  " Default number first, else the first one found (A30)
  LOOP AT gt_adr2 INTO DATA(ls_adr2) WHERE addrnumber = is_partner-adrnr.
    CASE ls_adr2-r3_user.
      WHEN '1'.
        IF cs_party-tel IS INITIAL OR ls_adr2-flgdefault = abap_true.
          cs_party-tel = ls_adr2-telnr_long.
        ENDIF.
      WHEN '3'.
        IF cs_party-mobile IS INITIAL OR ls_adr2-flgdefault = abap_true.
          cs_party-mobile = ls_adr2-telnr_long.
        ENDIF.
    ENDCASE.
  ENDLOOP.
  LOOP AT gt_adr6 INTO DATA(ls_adr6) WHERE addrnumber = is_partner-adrnr.
    IF cs_party-email IS INITIAL OR ls_adr6-flgdefault = abap_true.
      cs_party-email = ls_adr6-smtp_addr.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
FORM country_text USING    iv_land1 TYPE land1
                  CHANGING cv_text  TYPE string.

  CLEAR cv_text.
  IF iv_land1 IS INITIAL.
    RETURN.
  ENDIF.
  READ TABLE gt_t005t INTO DATA(ls_t005t) WITH TABLE KEY land1 = iv_land1.
  IF sy-subrc = 0 AND ls_t005t-landx IS NOT INITIAL.
    cv_text = to_upper( ls_t005t-landx ).
  ELSE.
    cv_text = iv_land1.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Part No. = SAP material number (logic B58 "from STEP no 53")
*&---------------------------------------------------------------------*
FORM matnr_of USING    iv_partno TYPE zsd_partno
              CHANGING cv_matnr  TYPE matnr.

  CLEAR cv_matnr.
  CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
    EXPORTING
      input        = iv_partno
    IMPORTING
      output       = cv_matnr
    EXCEPTIONS
      length_error = 1
      OTHERS       = 2.
  IF sy-subrc <> 0.
    CLEAR cv_matnr.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
FORM steuc_of USING    iv_hsn   TYPE csequence
              CHANGING cv_steuc TYPE steuc.
  cv_steuc = replace( val = condense( CONV string( iv_hsn ) ) sub = ` ` with = `` occ = 0 ).
ENDFORM.

*&---------------------------------------------------------------------*
*& Step 3: one Adobe Form per packing list number. An error is logged on
*& that packing list and the loop continues with the next one.
*&---------------------------------------------------------------------*
FORM output_forms.

  DATA lv_done TYPE abap_bool.
  DATA lt_err  TYPE bapiret2_t.

  LOOP AT gt_result ASSIGNING FIELD-SYMBOL(<ls_result>) WHERE status <> gc_st_error.

    CLEAR lt_err.
    PERFORM send_form USING <ls_result>-print CHANGING lv_done lt_err.

    IF lt_err IS NOT INITIAL.
      APPEND LINES OF lt_err TO <ls_result>-messages.
      <ls_result>-status = gc_st_error.
      <ls_result>-output = abap_false.
    ELSEIF lv_done = abap_true.
      <ls_result>-output = abap_true.
      PERFORM add_msg USING 'S' '065' <ls_result>-print-form <ls_result>-packno space space
                      CHANGING <ls_result>-messages.
    ELSE.
      PERFORM add_msg USING 'W' '064' <ls_result>-packno space space space CHANGING <ls_result>-messages.
      <ls_result>-status = gc_st_warning.
    ENDIF.

  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Call the Adobe Form for one packing list: print preview on LP01
*&---------------------------------------------------------------------*
FORM send_form USING    is_print    TYPE ty_print_data
               CHANGING cv_done     TYPE abap_bool
                        ct_messages TYPE bapiret2_t.

  DATA lv_fm      TYPE funcname.
  DATA ls_outpar  TYPE sfpoutputparams.
  DATA ls_docpar  TYPE sfpdocparams.
  DATA ls_formout TYPE fpformoutput.
  DATA lv_subrc   TYPE sysubrc.

  cv_done = abap_false.

  " Generated function module of the Adobe Form
  TRY.
      CALL FUNCTION 'FP_FUNCTION_MODULE_NAME'
        EXPORTING
          i_name     = is_print-form
        IMPORTING
          e_funcname = lv_fm.
    CATCH cx_fp_api.
      PERFORM add_msg USING 'E' '040' is_print-form space space space CHANGING ct_messages.
      RETURN.
  ENDTRY.

  " Print preview on LP01 - the user prints from the preview
  ls_outpar-nodialog = abap_true.
  ls_outpar-preview  = abap_true.
  ls_outpar-reqnew   = abap_true.
  ls_outpar-dest     = gc_dest.

  " One job per packing list: one preview each
  CALL FUNCTION 'FP_JOB_OPEN'
    CHANGING
      ie_outputparams = ls_outpar
    EXCEPTIONS
      cancel          = 1
      usage_error     = 2
      system_error    = 3
      internal_error  = 4
      OTHERS          = 5.
  IF sy-subrc = 1.
    RETURN.                                         " cancelled by the user
  ELSEIF sy-subrc <> 0.
    lv_subrc = sy-subrc.
    PERFORM add_msg USING 'E' '039' lv_subrc 'FP_JOB_OPEN' space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  ls_docpar-langu   = gc_langu.
  ls_docpar-country = gc_hsn_country.

  " Form interface ZSD_EPACK_AF_IF: IS_HEADER / IT_ITEM of THIS packing list
  CALL FUNCTION lv_fm
    EXPORTING
      /1bcdwb/docparams  = ls_docpar
      is_header          = is_print-header
      it_item            = is_print-items
    IMPORTING
      /1bcdwb/formoutput = ls_formout
    EXCEPTIONS
      usage_error        = 1
      system_error       = 2
      internal_error     = 3
      OTHERS             = 4.
  lv_subrc = sy-subrc.

  " The job is closed in every case
  CALL FUNCTION 'FP_JOB_CLOSE'
    EXCEPTIONS
      usage_error    = 1
      system_error   = 2
      internal_error = 3
      OTHERS         = 4.

  IF lv_subrc <> 0.
    PERFORM add_msg USING 'E' '039' lv_subrc is_print-form space space CHANGING ct_messages.
    RETURN.
  ENDIF.

  cv_done = abap_true.

ENDFORM.

*&---------------------------------------------------------------------*
*& Step 4: write every packing list whose form was produced to the log
*& table ZSD_EPACK_DATA - FS "Technical detail" layout: one row per line
*& item, the header fields repeated on every row, all invoice numbers of
*& the packing list in one text field VBRK_VBELN ("1000000348, ...").
*& A packing list uploaded again replaces its earlier rows.
*&---------------------------------------------------------------------*
FORM save_packing_lists.

  DATA lr_packno TYPE ty_packno_range.
  DATA lt_data   TYPE STANDARD TABLE OF zsd_epack_data WITH EMPTY KEY.
  DATA ls_data   TYPE zsd_epack_data.
  DATA lt_inv    TYPE ty_text_tab.
  DATA ls_totals TYPE ty_totals.
  DATA lv_ok     TYPE abap_bool.

  " Packing lists whose form was produced
  LOOP AT gt_result INTO DATA(ls_result) WHERE output = abap_true.
    APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_result-packno ) TO lr_packno.
  ENDLOOP.

  IF lr_packno IS INITIAL.
    RETURN.
  ENDIF.

  LOOP AT lr_packno INTO DATA(ls_packno).
    READ TABLE gt_packing_lists INTO DATA(ls_pl) WITH TABLE KEY packno = ls_packno-low.
    IF sy-subrc <> 0.
      CONTINUE.
    ENDIF.

    " All invoice numbers in one field
    CLEAR lt_inv.
    LOOP AT ls_pl-invoices INTO DATA(lv_vbeln).
      APPEND |{ lv_vbeln ALPHA = OUT }| TO lt_inv.
    ENDLOOP.

    " Total Articles (logic B52) - calculated, the template has no column
    PERFORM calc_totals USING ls_pl-items CHANGING ls_totals.

    " Header fields of the packing list (same on every row)
    CLEAR ls_data.
    MOVE-CORRESPONDING ls_pl-header TO ls_data.
    ls_data-vbrk_vbeln = concat_lines_of( table = lt_inv sep = `, ` ).
    ls_data-zsd_tart   = ls_totals-articles.
    ls_data-zsd_meins  = gc_uom_qty.
    ls_data-zsd_gewei  = gc_uom_weight.

    " One row per line item, exactly as uploaded
    LOOP AT ls_pl-items INTO DATA(ls_item).
      ls_data-zsd_sn      = ls_item-zsd_sn.
      ls_data-zsd_inv_sn  = ls_item-zsd_inv_sn.
      ls_data-zsd_artno   = ls_item-zsd_artno.
      ls_data-zsd_std     = ls_item-zsd_std.
      ls_data-zsd_art     = ls_item-zsd_art.
      ls_data-zsd_partno  = ls_item-zsd_partno.
      ls_data-zsd_matdesc = ls_item-zsd_matdesc.
      ls_data-zsd_qty     = ls_item-zsd_qty.
      ls_data-zsd_gwt     = ls_item-zsd_gwt.
      ls_data-zsd_nwt     = ls_item-zsd_nwt.
      ls_data-zsd_hsn     = ls_item-zsd_hsn.
      APPEND ls_data TO lt_data.
    ENDLOOP.
  ENDLOOP.

  " Replace earlier rows of these packing lists, write the new ones
  DELETE FROM zsd_epack_data WHERE zsd_packno IN @lr_packno.
  INSERT zsd_epack_data FROM TABLE @lt_data.
  lv_ok = xsdbool( sy-subrc = 0 ).

  LOOP AT lr_packno INTO ls_packno.
    READ TABLE gt_result ASSIGNING FIELD-SYMBOL(<ls_result>) WITH TABLE KEY packno = ls_packno-low.
    IF lv_ok = abap_true.
      <ls_result>-saved = abap_true.
      PERFORM add_msg USING 'S' '032' ls_packno-low <ls_result>-item_count <ls_result>-inv_count space
                      CHANGING <ls_result>-messages.
    ELSE.
      PERFORM add_msg USING 'E' '052' ls_packno-low space space space CHANGING <ls_result>-messages.
    ENDIF.
  ENDLOOP.

  IF lv_ok = abap_false.
    " Nothing half-written: undo the delete as well
    ROLLBACK WORK.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& One COMMIT WORK for the log table
*&---------------------------------------------------------------------*
FORM finish_db.

  COMMIT WORK AND WAIT.

ENDFORM.

*&---------------------------------------------------------------------*
*& Overall status message
*&---------------------------------------------------------------------*
FORM final_message.

  DATA lv_total  TYPE i.
  DATA lv_output TYPE i.
  DATA lv_saved  TYPE i.
  DATA lv_error  TYPE i.

  lv_total = lines( gt_result ).
  LOOP AT gt_result INTO DATA(ls_res).
    IF ls_res-output = abap_true.
      lv_output = lv_output + 1.
    ENDIF.
    IF ls_res-saved = abap_true.
      lv_saved = lv_saved + 1.
    ENDIF.
    IF ls_res-status = gc_st_error.
      lv_error = lv_error + 1.
    ENDIF.
  ENDLOOP.

  IF lv_total = 0.
    MESSAGE s059(zsd_epack) DISPLAY LIKE 'E'.
  ELSEIF lv_error > 0.
    MESSAGE s066(zsd_epack) WITH lv_output lv_total lv_saved lv_error DISPLAY LIKE 'W'.
  ELSE.
    MESSAGE s066(zsd_epack) WITH lv_output lv_total lv_saved lv_error.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Result list: per packing list status, form, counts, output / saved
*& flags and every message; file-level messages first
*&---------------------------------------------------------------------*
FORM show_result.

  TYPES: BEGIN OF ty_line,
           icon    TYPE icon_d,
           packno  TYPE zsd_packno,
           status  TYPE char1,
           form    TYPE fpname,
           inv     TYPE i,
           items   TYPE i,
           output  TYPE xfeld,
           saved   TYPE xfeld,
           type    TYPE bapi_mtype,
           number  TYPE symsgno,
           message TYPE bapi_msg,
         END OF ty_line.
  DATA lt_line  TYPE STANDARD TABLE OF ty_line WITH EMPTY KEY.
  DATA ls_base  TYPE ty_line.
  DATA ls_line  TYPE ty_line.
  DATA lo_alv   TYPE REF TO cl_salv_table.
  DATA lx_salv  TYPE REF TO cx_salv_error.

  LOOP AT gt_file_messages INTO DATA(ls_file).
    APPEND VALUE #( icon = icon_led_red type = ls_file-type number = ls_file-number
                    message = ls_file-message ) TO lt_line.
  ENDLOOP.

  LOOP AT gt_result INTO DATA(ls_result).
    CLEAR ls_base.
    ls_base-icon   = SWITCH #( ls_result-status WHEN gc_st_error   THEN icon_led_red
                                                WHEN gc_st_warning THEN icon_led_yellow
                                                ELSE icon_led_green ).
    ls_base-packno = ls_result-packno.
    ls_base-status = ls_result-status.
    ls_base-form   = ls_result-print-form.
    ls_base-inv    = ls_result-inv_count.
    ls_base-items  = ls_result-item_count.
    ls_base-output = ls_result-output.
    ls_base-saved  = ls_result-saved.
    IF ls_result-messages IS INITIAL.
      APPEND ls_base TO lt_line.
    ENDIF.
    LOOP AT ls_result-messages INTO DATA(ls_msg).
      ls_line = ls_base.
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
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv
                              CHANGING  t_table      = lt_line ).
      lo_alv->get_functions( )->set_all( abap_true ).
      lo_alv->get_columns( )->set_optimize( abap_true ).
      lo_alv->get_columns( )->get_column( 'INV' )->set_short_text( 'Invoices' ).
      lo_alv->get_columns( )->get_column( 'ITEMS' )->set_short_text( 'Items' ).
      lo_alv->get_columns( )->get_column( 'OUTPUT' )->set_short_text( 'Output' ).
      lo_alv->get_columns( )->get_column( 'SAVED' )->set_short_text( 'Saved' ).
      lo_alv->get_sorts( )->add_sort( columnname = 'PACKNO' ).
      lo_alv->get_display_settings( )->set_list_header(
        CONV #( |Export Packing List - { lines( gt_result ) } packing list(s)| ) ).
      lo_alv->display( ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE s063(zsd_epack) DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& F4 help: Excel file
*&---------------------------------------------------------------------*
FORM f4_file CHANGING cv_file TYPE rlgrap-filename.

  DATA lt_files TYPE filetable.
  DATA lv_rc    TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING  file_filter    = 'Excel (*.xlsx)|*.xlsx'
               multiselection = abap_false
    CHANGING   file_table     = lt_files
               rc             = lv_rc
    EXCEPTIONS OTHERS         = 1 ).
  IF sy-subrc = 0 AND lv_rc = 1.
    READ TABLE lt_files INTO DATA(ls_file) INDEX 1.
    cv_file = ls_file-filename.
  ENDIF.

ENDFORM.

