"! <p class="shorttext synchronized">Export Packing List - Excel parser</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Reads the FS template "Z table Format1.xlsx" (sheet 1, unchanged layout)
"! and splits it into PACKING LISTS. No database access - all SAP checks
"! are done in bulk by ZCL_SD_EPACK_DATA afterwards.
"!
"! FILE LAYOUT (columns as in the FS template, A = 1)
"!   row 1       column headings (checked loosely)
"!   A-K         line item: Sl.No., Sr.No. INV, Article No., STD, Articles,
"!               Part No., Product Description, Quantity, Gross/Net Wt., H.S Code
"!   L           SAP Invoice Number - one per row and/or several per cell
"!   M           Packing list no
"!   N-AK        header data of the packing list (logic sheet "Pass
"!               ZSD_PACKNO get ...")
"!   AL          optional Package Type (not in the FS template, see A09)
"!
"! SEVERAL PACKING LISTS PER FILE (ASSUMPTION A33)
"!   A row with "Packing list no" filled starts / continues that packing
"!   list; rows with the column empty belong to the packing list of the
"!   nearest row above. Each packing list gets its own header, invoices,
"!   items and messages - nothing is shared between packing lists.
"!
"! INVOICE NUMBERS
"!   All invoice numbers of the rows of a packing list belong to that
"!   packing list (FS: "Input will be tax invoices (multiple tax invoices)");
"!   repeated numbers are counted once. An invoice that appears under two
"!   packing lists of the same file is an error for both (A01).
CLASS zcl_sd_epack_upload DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    TYPES tt_cells TYPE STANDARD TABLE OF string WITH EMPTY KEY.
    TYPES:
      BEGIN OF ty_row,
        row   TYPE i,
        cells TYPE tt_cells,
      END OF ty_row,
      tt_rows TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

    "! Read a file from the presentation server
    "! @raising zcx_sd_epack | file not readable
    CLASS-METHODS read_file
      IMPORTING iv_path           TYPE string
      RETURNING VALUE(rv_xstring) TYPE xstring
      RAISING   zcx_sd_epack.

    "! Excel content -> packing lists
    "! @parameter et_messages | file-level messages (structure, rows
    "!                          without packing list number)
    METHODS parse
      IMPORTING iv_xstring       TYPE xstring
                iv_filename      TYPE string
      EXPORTING et_packing_lists TYPE zif_sd_epack=>tt_packing_list
                et_messages      TYPE bapiret2_t.

    "! Worksheet rows (row 1 = headings) -> packing lists.
    "! Public for ABAP Unit.
    METHODS parse_rows
      IMPORTING it_rows          TYPE tt_rows
                iv_filename      TYPE string OPTIONAL
      EXPORTING et_packing_lists TYPE zif_sd_epack=>tt_packing_list
                et_messages      TYPE bapiret2_t.

  PRIVATE SECTION.

    "! Header column -> field of ZSD_EPACK_HDR
    TYPES:
      BEGIN OF ty_map,
        col    TYPE i,
        field  TYPE fieldname,
        "! C char / T multi-line text / D date / A amount
        kind   TYPE char1,
        "! maximum length, 0 = unlimited
        maxlen TYPE i,
        label  TYPE string,
      END OF ty_map,
      tt_map TYPE STANDARD TABLE OF ty_map WITH EMPTY KEY.

    TYPES:
      BEGIN OF ty_pl_rows,
        packno TYPE zsd_packno,
        rows   TYPE tt_rows,
      END OF ty_pl_rows,
      tt_pl_rows TYPE STANDARD TABLE OF ty_pl_rows WITH EMPTY KEY.

    "! Column numbers of the FS template (A = 1)
    CONSTANTS:
      BEGIN OF gc_col,
        sn       TYPE i VALUE 1,
        inv_sn   TYPE i VALUE 2,
        artno    TYPE i VALUE 3,
        std      TYPE i VALUE 4,
        art      TYPE i VALUE 5,
        partno   TYPE i VALUE 6,
        matdesc  TYPE i VALUE 7,
        qty      TYPE i VALUE 8,
        gwt      TYPE i VALUE 9,
        nwt      TYPE i VALUE 10,
        hsn      TYPE i VALUE 11,
        invoice  TYPE i VALUE 12,
        packno   TYPE i VALUE 13,
        pkg_type TYPE i VALUE 38,
      END OF gc_col.

    "! Messages go to the file or to the packing list being processed
    DATA mr_messages TYPE REF TO bapiret2_t.

    METHODS add
      IMPORTING iv_type TYPE bapi_mtype DEFAULT 'E'
                iv_no   TYPE symsgno
                iv_v1   TYPE simple OPTIONAL
                iv_v2   TYPE simple OPTIONAL
                iv_v3   TYPE simple OPTIONAL
                iv_v4   TYPE simple OPTIONAL.

    METHODS read_worksheet
      IMPORTING iv_xstring     TYPE xstring
                iv_filename    TYPE string
      RETURNING VALUE(rt_rows) TYPE tt_rows
      RAISING   zcx_sd_epack.

    METHODS check_template
      IMPORTING is_row TYPE ty_row.

    "! Split rows by packing list number (carry-forward of column M)
    METHODS split_by_packing_list
      IMPORTING it_rows        TYPE tt_rows
      RETURNING VALUE(rt_pl)   TYPE tt_pl_rows.

    METHODS map_header
      IMPORTING it_rows TYPE tt_rows
      CHANGING  cs_pl   TYPE zif_sd_epack=>ts_packing_list.

    METHODS map_lines
      IMPORTING it_rows TYPE tt_rows
      CHANGING  cs_pl   TYPE zif_sd_epack=>ts_packing_list.

    METHODS cell
      IMPORTING is_row         TYPE ty_row
                iv_col         TYPE i
      RETURNING VALUE(rv_text) TYPE string.

    METHODS number
      IMPORTING is_row          TYPE ty_row
                iv_col          TYPE i
                iv_integer      TYPE abap_bool DEFAULT abap_false
      RETURNING VALUE(rv_value) TYPE decfloat34.

    METHODS check_length
      IMPORTING iv_row  TYPE i
                iv_col  TYPE i
                iv_text TYPE string
                iv_max  TYPE i.

    CLASS-METHODS header_map
      RETURNING VALUE(rt_map) TYPE tt_map.

    CLASS-METHODS column_name
      IMPORTING iv_col         TYPE i
      RETURNING VALUE(rv_name) TYPE string.

ENDCLASS.


CLASS zcl_sd_epack_upload IMPLEMENTATION.

  METHOD read_file.

    DATA lt_solix  TYPE solix_tab.
    DATA lv_length TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING  filename   = iv_path
                 filetype   = 'BIN'
      IMPORTING  filelength = lv_length
      CHANGING   data_tab   = lt_solix
      EXCEPTIONS OTHERS     = 1 ).
    IF sy-subrc <> 0.
      zcx_sd_epack=>raise( iv_msgno = '015' iv_v1 = iv_path iv_v2 = 'GUI_UPLOAD' ).
    ENDIF.

    IF lv_length = 0.
      zcx_sd_epack=>raise( iv_msgno = '036' ).
    ENDIF.

    rv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_solix
                                                   iv_size  = lv_length ).

  ENDMETHOD.


  METHOD parse.

    CLEAR: et_packing_lists, et_messages.

    TRY.
        DATA(lt_rows) = read_worksheet( iv_xstring  = iv_xstring
                                        iv_filename = iv_filename ).
      CATCH zcx_sd_epack INTO DATA(lx_error).
        et_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '015'
                                                          iv_v1 = iv_filename
                                                          iv_v2 = lx_error->get_text( ) ) ) ).
        RETURN.
    ENDTRY.

    parse_rows( EXPORTING it_rows          = lt_rows
                          iv_filename      = iv_filename
                IMPORTING et_packing_lists = et_packing_lists
                          et_messages      = et_messages ).

  ENDMETHOD.


  METHOD parse_rows.

    DATA lt_file_msg TYPE bapiret2_t.

    CLEAR: et_packing_lists, et_messages.
    mr_messages = REF #( lt_file_msg ).

    IF lines( it_rows ) < 2.
      add( iv_no = '036' ).
      et_messages = lt_file_msg.
      RETURN.
    ENDIF.

    check_template( it_rows[ 1 ] ).
    IF zcl_sd_epack_rules=>has_errors( lt_file_msg ).
      et_messages = lt_file_msg.
      RETURN.
    ENDIF.

    DATA(lt_data_rows) = it_rows.
    DELETE lt_data_rows INDEX 1.

    DATA(lt_split) = split_by_packing_list( lt_data_rows ).
    IF lt_split IS INITIAL.
      add( iv_no = '059' ).
      et_messages = lt_file_msg.
      RETURN.
    ENDIF.

    " --- One packing list at a time: header, invoices, items, checks -----
    LOOP AT lt_split INTO DATA(ls_split).

      DATA(ls_pl) = VALUE zif_sd_epack=>ts_packing_list( packno = ls_split-packno ).
      ls_pl-header-zsd_packno   = ls_split-packno.
      ls_pl-header-zsd_filename = iv_filename.
      ls_pl-rows = |{ ls_split-rows[ 1 ]-row }-{ ls_split-rows[ lines( ls_split-rows ) ]-row }|.

      mr_messages = REF #( ls_pl-messages ).

      map_header( EXPORTING it_rows = ls_split-rows CHANGING cs_pl = ls_pl ).
      map_lines(  EXPORTING it_rows = ls_split-rows CHANGING cs_pl = ls_pl ).

      " FS 2.1 weight rule and the quantity plausibility per packing list
      APPEND LINES OF zcl_sd_epack_rules=>check_weights( ls_pl-items )    TO ls_pl-messages.
      APPEND LINES OF zcl_sd_epack_rules=>check_quantities( ls_pl-items ) TO ls_pl-messages.

      INSERT ls_pl INTO TABLE et_packing_lists.
    ENDLOOP.

    " --- Cross-check: one invoice may not appear in two packing lists ------
    TYPES: BEGIN OF ty_inv_pl,
             vbeln  TYPE vbeln_vf,
             packno TYPE zsd_packno,
           END OF ty_inv_pl.
    DATA lt_inv_pl TYPE SORTED TABLE OF ty_inv_pl WITH NON-UNIQUE KEY vbeln.

    LOOP AT et_packing_lists INTO ls_pl.
      LOOP AT ls_pl-invoices INTO DATA(lv_vbeln).
        INSERT VALUE #( vbeln = lv_vbeln packno = ls_pl-packno ) INTO TABLE lt_inv_pl.
      ENDLOOP.
    ENDLOOP.

    LOOP AT lt_inv_pl INTO DATA(ls_inv_pl).
      LOOP AT lt_inv_pl INTO DATA(ls_other)
           WHERE vbeln = ls_inv_pl-vbeln AND packno <> ls_inv_pl-packno.
        ASSIGN et_packing_lists[ packno = ls_inv_pl-packno ] TO FIELD-SYMBOL(<ls_pl>).
        IF sy-subrc = 0.
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '058'
                                          iv_v1 = |{ ls_inv_pl-vbeln ALPHA = OUT }|
                                          iv_v2 = ls_other-packno ) TO <ls_pl>-messages.
        ENDIF.
      ENDLOOP.
    ENDLOOP.

    et_messages = lt_file_msg.

  ENDMETHOD.


  METHOD add.
    APPEND zcl_sd_epack_rules=>msg( iv_type = iv_type iv_no = iv_no
                                    iv_v1 = iv_v1 iv_v2 = iv_v2
                                    iv_v3 = iv_v3 iv_v4 = iv_v4 ) TO mr_messages->*.
  ENDMETHOD.


  METHOD read_worksheet.

    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

    " CL_FDT_XL_SPREADSHEET reads .xlsx in the application server memory
    " (no SAP GUI / OLE automation). ⚠ SAP-VERIFY: not a released API; on
    " newer S/4HANA releases XCO_CP_XLSX is the released alternative (A21).
    TRY.
        DATA(lo_excel) = NEW cl_fdt_xl_spreadsheet( document_name = iv_filename
                                                    xdocument     = iv_xstring ).
        lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
          IMPORTING worksheet_names = DATA(lt_sheets) ).
        IF lt_sheets IS INITIAL.
          RETURN.
        ENDIF.
        DATA(lr_sheet) = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lt_sheets[ 1 ] ).
      CATCH cx_fdt_excel_core INTO DATA(lx_excel).
        zcx_sd_epack=>raise( iv_msgno = '015' iv_v1 = iv_filename iv_v2 = lx_excel->get_text( ) ).
    ENDTRY.

    ASSIGN lr_sheet->* TO <lt_sheet>.
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    LOOP AT <lt_sheet> ASSIGNING FIELD-SYMBOL(<ls_line>).
      DATA(ls_row) = VALUE ty_row( row = sy-tabix ).
      DO.
        ASSIGN COMPONENT sy-index OF STRUCTURE <ls_line> TO FIELD-SYMBOL(<lv_cell>).
        IF sy-subrc <> 0.
          EXIT.
        ENDIF.
        APPEND condense( val = CONV string( <lv_cell> ) del = ` ` ) TO ls_row-cells.
      ENDDO.
      " Skip completely empty rows
      LOOP AT ls_row-cells TRANSPORTING NO FIELDS WHERE table_line IS NOT INITIAL.
        APPEND ls_row TO rt_rows.
        EXIT.
      ENDLOOP.
    ENDLOOP.

  ENDMETHOD.


  METHOD check_template.

    " Loose check: column A must be "Sl.No." and column M "Packing list no"
    " - catches files in a wrong or shifted layout.
    DATA(lv_a) = to_upper( cell( is_row = is_row iv_col = gc_col-sn ) ).
    DATA(lv_m) = to_upper( cell( is_row = is_row iv_col = gc_col-packno ) ).

    IF lv_a NS 'SL'.
      add( iv_no = '016' iv_v1 = 'A' iv_v2 = cell( is_row = is_row iv_col = gc_col-sn ) ).
    ENDIF.
    IF lv_m NS 'PACKING'.
      add( iv_no = '016' iv_v1 = 'M' iv_v2 = cell( is_row = is_row iv_col = gc_col-packno ) ).
    ENDIF.

  ENDMETHOD.


  METHOD split_by_packing_list.

    DATA lv_current TYPE zsd_packno.

    LOOP AT it_rows INTO DATA(ls_row).

      DATA(lv_packno) = cell( is_row = ls_row iv_col = gc_col-packno ).
      IF lv_packno IS NOT INITIAL.
        IF strlen( lv_packno ) > 10.
          add( iv_no = '022' iv_v1 = ls_row-row iv_v2 = 'M' iv_v3 = 10 ).
          CONTINUE.
        ENDIF.
        lv_current = |{ lv_packno ALPHA = IN WIDTH = 10 }|.
      ELSEIF lv_current IS INITIAL.
        " Data above the first packing list number cannot be assigned
        add( iv_no = '055' iv_v1 = ls_row-row ).
        CONTINUE.
      ENDIF.

      ASSIGN rt_pl[ packno = lv_current ] TO FIELD-SYMBOL(<ls_pl>).
      IF sy-subrc <> 0.
        APPEND VALUE #( packno = lv_current ) TO rt_pl ASSIGNING <ls_pl>.
      ENDIF.
      APPEND ls_row TO <ls_pl>-rows.

    ENDLOOP.

  ENDMETHOD.


  METHOD map_header.

    TYPES: BEGIN OF ty_first,
             col   TYPE i,
             row   TYPE i,
             value TYPE string,
           END OF ty_first.
    DATA lt_first      TYPE SORTED TABLE OF ty_first WITH UNIQUE KEY col.
    DATA lv_date       TYPE d.
    DATA lv_num        TYPE decfloat34.
    DATA lv_inv_found  TYPE i.

    DATA(lt_map) = header_map( ).

    LOOP AT it_rows INTO DATA(ls_row).

      " Invoices: every row of the packing list may carry one or more
      DATA(lv_inv) = cell( is_row = ls_row iv_col = gc_col-invoice ).
      IF lv_inv IS NOT INITIAL.
        zcl_sd_epack_rules=>split_invoices( EXPORTING iv_text    = lv_inv
                                            IMPORTING et_vbeln   = DATA(lt_vbeln)
                                                      et_invalid = DATA(lt_invalid) ).
        lv_inv_found += lines( lt_vbeln ).
        " Sorted unique table - a repeated invoice is kept once
        INSERT LINES OF lt_vbeln INTO TABLE cs_pl-invoices.
        LOOP AT lt_invalid INTO DATA(lv_invalid).
          add( iv_no = '018' iv_v1 = ls_row-row iv_v2 = lv_invalid iv_v3 = 'L' ).
        ENDLOOP.
      ENDIF.

      LOOP AT lt_map INTO DATA(ls_map).
        DATA(lv_value) = cell( is_row = ls_row iv_col = ls_map-col ).
        IF lv_value IS INITIAL.
          CONTINUE.
        ENDIF.

        " The same value in several rows is fine, a different one is not
        READ TABLE lt_first INTO DATA(ls_first) WITH TABLE KEY col = ls_map-col.
        IF sy-subrc = 0.
          IF ls_first-value <> lv_value.
            add( iv_no = '020' iv_v1 = ls_row-row iv_v2 = ls_map-label iv_v3 = ls_first-row ).
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
            IF zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = lv_value
                                                   IMPORTING ev_date = lv_date ) = abap_true.
              <lv_field> = lv_date.
            ELSE.
              add( iv_no = '019' iv_v1 = ls_row-row iv_v2 = lv_value iv_v3 = column_name( ls_map-col ) ).
            ENDIF.
          WHEN 'A'.
            IF zcl_sd_epack_rules=>try_parse_number( EXPORTING iv_text  = lv_value
                                                     IMPORTING ev_value = lv_num ) = abap_true.
              <lv_field> = lv_num.
            ELSE.
              add( iv_no = '018' iv_v1 = ls_row-row iv_v2 = lv_value iv_v3 = column_name( ls_map-col ) ).
            ENDIF.
          WHEN OTHERS.
            " Multi-line cells (Alt+Enter) keep their line breaks (kind T)
            check_length( iv_row = ls_row-row iv_col = ls_map-col iv_text = lv_value iv_max = ls_map-maxlen ).
            <lv_field> = lv_value.
        ENDCASE.
      ENDLOOP.

    ENDLOOP.

    DATA(lv_repeated) = lv_inv_found - lines( cs_pl-invoices ).
    IF lv_repeated > 0.
      add( iv_type = 'I' iv_no = '056' iv_v1 = lv_repeated ).
    ENDIF.

    " Mandatory per packing list
    DATA(lv_first_row) = it_rows[ 1 ]-row.
    IF cs_pl-header-zsd_packdt IS INITIAL.
      add( iv_no = '017' iv_v1 = lv_first_row iv_v2 = 'Packing list date' ).
    ENDIF.
    IF cs_pl-invoices IS INITIAL.
      add( iv_no = '017' iv_v1 = lv_first_row iv_v2 = 'SAP Invoice Number' ).
    ENDIF.

    " Notify parties must be filled from 1 upwards (A10)
    DATA lv_gap TYPE i.
    DO zif_sd_epack=>gc_max_np TIMES.
      ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE cs_pl-header TO FIELD-SYMBOL(<lv_np>).
      IF <lv_np> IS INITIAL.
        IF lv_gap IS INITIAL.
          lv_gap = sy-index.
        ENDIF.
      ELSEIF lv_gap IS NOT INITIAL.
        add( iv_no = '031' iv_v1 = sy-index iv_v2 = lv_gap ).
      ENDIF.
    ENDDO.

  ENDMETHOD.


  METHOD map_lines.

    DATA lt_lines   TYPE zif_sd_epack=>tt_upload_line.
    DATA lv_ignored TYPE i.

    LOOP AT it_rows INTO DATA(ls_row).

      " A line item needs a Sl.No.; other rows (total row of the template,
      " header-only rows) are not line items
      IF cell( is_row = ls_row iv_col = gc_col-sn ) IS INITIAL.
        IF cell( is_row = ls_row iv_col = gc_col-partno ) IS NOT INITIAL
           OR cell( is_row = ls_row iv_col = gc_col-qty ) IS NOT INITIAL.
          lv_ignored += 1.
        ENDIF.
        CONTINUE.
      ENDIF.

      DATA(ls_line) = VALUE zif_sd_epack=>ts_upload_line( row = ls_row-row ).

      ls_line-sn     = number( is_row = ls_row iv_col = gc_col-sn     iv_integer = abap_true ).
      ls_line-inv_sn = number( is_row = ls_row iv_col = gc_col-inv_sn iv_integer = abap_true ).
      ls_line-std    = number( is_row = ls_row iv_col = gc_col-std    iv_integer = abap_true ).
      ls_line-art    = number( is_row = ls_row iv_col = gc_col-art    iv_integer = abap_true ).
      ls_line-qty    = round( val = number( is_row = ls_row iv_col = gc_col-qty ) dec = 3 ).
      ls_line-gwt    = round( val = number( is_row = ls_row iv_col = gc_col-gwt ) dec = 3 ).
      ls_line-nwt    = round( val = number( is_row = ls_row iv_col = gc_col-nwt ) dec = 3 ).

      " Empty cell = merged cell, i.e. part of the group above (A06)
      ls_line-art_filled = xsdbool( cell( is_row = ls_row iv_col = gc_col-art ) IS NOT INITIAL ).
      ls_line-gwt_filled = xsdbool( cell( is_row = ls_row iv_col = gc_col-gwt ) IS NOT INITIAL ).
      ls_line-nwt_filled = xsdbool( cell( is_row = ls_row iv_col = gc_col-nwt ) IS NOT INITIAL ).

      DATA(lv_artno)   = cell( is_row = ls_row iv_col = gc_col-artno ).
      DATA(lv_partno)  = cell( is_row = ls_row iv_col = gc_col-partno ).
      DATA(lv_matdesc) = cell( is_row = ls_row iv_col = gc_col-matdesc ).
      DATA(lv_hsn)     = cell( is_row = ls_row iv_col = gc_col-hsn ).
      DATA(lv_pkg)     = cell( is_row = ls_row iv_col = gc_col-pkg_type ).

      check_length( iv_row = ls_row-row iv_col = gc_col-artno    iv_text = lv_artno   iv_max = 15 ).
      check_length( iv_row = ls_row-row iv_col = gc_col-partno   iv_text = lv_partno  iv_max = 40 ).
      check_length( iv_row = ls_row-row iv_col = gc_col-matdesc  iv_text = lv_matdesc iv_max = 40 ).
      check_length( iv_row = ls_row-row iv_col = gc_col-hsn      iv_text = lv_hsn     iv_max = 17 ).
      check_length( iv_row = ls_row-row iv_col = gc_col-pkg_type iv_text = lv_pkg     iv_max = 20 ).

      ls_line-artno    = lv_artno.
      ls_line-partno   = to_upper( lv_partno ).
      ls_line-matdesc  = lv_matdesc.
      ls_line-hsn      = lv_hsn.
      ls_line-pkg_type = lv_pkg.

      IF ls_line-partno IS INITIAL.
        add( iv_no = '017' iv_v1 = ls_row-row iv_v2 = 'Part No.' ).
      ENDIF.
      IF cell( is_row = ls_row iv_col = gc_col-qty ) IS INITIAL.
        add( iv_no = '017' iv_v1 = ls_row-row iv_v2 = 'Quantity' ).
      ENDIF.

      " Exact duplicate of an earlier row (same Sl.No. and same content):
      " counted once (A38). Same Sl.No. with different content is an
      " error raised by BUILD_ITEMS.
      LOOP AT lt_lines INTO DATA(ls_prev) WHERE sn = ls_line-sn.
        DATA(ls_cmp) = ls_line.
        ls_cmp-row = ls_prev-row.
        IF ls_cmp = ls_prev.
          add( iv_type = 'W' iv_no = '057' iv_v1 = ls_row-row iv_v2 = ls_prev-row ).
          CLEAR ls_line.
        ENDIF.
        EXIT.
      ENDLOOP.
      IF ls_line IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND ls_line TO lt_lines.
    ENDLOOP.

    IF lv_ignored > 0.
      add( iv_type = 'I' iv_no = '050' iv_v1 = lv_ignored ).
    ENDIF.

    IF lt_lines IS INITIAL.
      add( iv_no = '036' ).
      RETURN.
    ENDIF.

    zcl_sd_epack_rules=>build_items( EXPORTING iv_packno   = cs_pl-packno
                                               it_lines    = lt_lines
                                     IMPORTING et_items    = cs_pl-items
                                     CHANGING  ct_messages = mr_messages->* ).

  ENDMETHOD.


  METHOD cell.
    rv_text = VALUE #( is_row-cells[ iv_col ] OPTIONAL ).
  ENDMETHOD.


  METHOD number.

    DATA lv_value TYPE decfloat34.

    DATA(lv_text) = cell( is_row = is_row iv_col = iv_col ).
    IF lv_text IS INITIAL.
      RETURN.
    ENDIF.

    IF zcl_sd_epack_rules=>try_parse_number( EXPORTING iv_text  = lv_text
                                             IMPORTING ev_value = lv_value ) = abap_false
       OR lv_value < 0
       OR ( iv_integer = abap_true AND ( frac( lv_value ) <> 0 OR lv_value > 999999 ) ).
      add( iv_no = '018' iv_v1 = is_row-row iv_v2 = lv_text iv_v3 = column_name( iv_col ) ).
      RETURN.
    ENDIF.

    rv_value = lv_value.

  ENDMETHOD.


  METHOD check_length.

    IF iv_max > 0 AND strlen( iv_text ) > iv_max.
      add( iv_no = '022' iv_v1 = iv_row iv_v2 = column_name( iv_col ) iv_v3 = iv_max ).
    ENDIF.

  ENDMETHOD.


  METHOD header_map.

    " Column order and lengths of "Z table Format1.xlsx" / Technical detail.
    " Column O "Exporter ReF" is NOT mapped: Technical detail marks
    " ZSD_EXPREF "Not Req - This data will be fetched from SAP System"
    " (logic sheet B11). Multi-line fields are wider than in the FS because
    " the sample outputs do not fit into the FS lengths (A19).
    rt_map = VALUE #(
      ( col = 14 field = 'ZSD_PACKDT'    kind = 'D'               label = `Packing list date` )
      ( col = 16 field = 'ZSD_PARTY_REF' kind = 'C' maxlen = 200  label = `Buyer's Reference No.` )
      ( col = 17 field = 'ZSD_REF_DT'    kind = 'D'               label = `Buyer's Reference date` )
      ( col = 18 field = 'ZSD_CNTY_ORGN' kind = 'C' maxlen = 20   label = `Country of origin` )
      ( col = 19 field = 'ZSD_NP1'       kind = 'T' maxlen = 1333 label = `Notifier Party1` )
      ( col = 20 field = 'ZSD_NP2'       kind = 'T' maxlen = 1333 label = `Notifier Party2` )
      ( col = 21 field = 'ZSD_NP3'       kind = 'T' maxlen = 1333 label = `Notifier Party3` )
      ( col = 22 field = 'ZSD_NP4'       kind = 'T' maxlen = 1333 label = `Notifier Party4` )
      ( col = 23 field = 'ZSD_NP5'       kind = 'T' maxlen = 1333 label = `Notifier Party5` )
      ( col = 24 field = 'ZSD_PREC'      kind = 'C' maxlen = 30   label = `Pre-Carriage` )
      ( col = 25 field = 'ZSD_REC_PC'    kind = 'C' maxlen = 30   label = `Place of Receipt by Pre-carrier` )
      ( col = 26 field = 'ZSD_POL'       kind = 'C' maxlen = 30   label = `Port of Loading` )
      ( col = 27 field = 'ZSD_POD'       kind = 'C' maxlen = 30   label = `Port of Discharge` )
      ( col = 28 field = 'ZSD_PLD'       kind = 'C' maxlen = 30   label = `Place of Delivery` )
      ( col = 29 field = 'ZSD_MARK'      kind = 'T' maxlen = 255  label = `Marks & Nos.` )
      ( col = 30 field = 'ZSD_VSL_FLT'   kind = 'C' maxlen = 30   label = `Vessel / Flight No.` )
      ( col = 31 field = 'ZSD_CONT'      kind = 'T' maxlen = 1333 label = `Container No` )
      ( col = 32 field = 'ZSD_NOPACK'    kind = 'T' maxlen = 255  label = `No. of Packages` )
      ( col = 33 field = 'ZSD_HSN_DESC'  kind = 'T' maxlen = 1333 label = `Description of goods` )
      ( col = 34 field = 'ZSD_PL_DECL'   kind = 'T' maxlen = 0    label = `Declaration of Packing list` )
      ( col = 35 field = 'ZSD_ADV_AMT'   kind = 'A'               label = `Advance amount received` )
      ( col = 36 field = 'ZSD_CMMT'      kind = 'T' maxlen = 1333 label = `Commercial terms of invoice` )
      ( col = 37 field = 'ZSD_INV_DECL'  kind = 'T' maxlen = 0    label = `Declaration of Invoice` ) ).

  ENDMETHOD.


  METHOD column_name.

    " 1 -> A, 27 -> AA, 38 -> AL
    DATA(lv_col) = iv_col.
    WHILE lv_col > 0.
      DATA(lv_rem) = ( lv_col - 1 ) MOD 26.
      rv_name = |{ sy-abcde+lv_rem(1) }{ rv_name }|.
      lv_col = ( lv_col - 1 - lv_rem ) DIV 26.
    ENDWHILE.

  ENDMETHOD.

ENDCLASS.
