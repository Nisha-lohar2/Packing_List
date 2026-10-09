"! <p class="shorttext synchronized">Export Packing List - Excel upload</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! FS 1 / 2.1: "Any information that is not available in SAP should be
"! maintained by users through an Excel upload into a custom Z-table."
"! The FS does not specify the upload itself (clarification CL-11);
"! this class implements ASSUMPTIONS A09, A10, A19 and A21-A23.
"!
"! FILE LAYOUT = the FS template "Z table Format1.xlsx", sheet 1, unchanged:
"!   row 1        column headings (checked loosely)
"!   rows 2..n    one row per line item - columns A-K
"!                header data (L-AK) may be entered in any row, normally
"!                the first; a value repeated in later rows must be equal
"!   column L     SAP invoice number(s): one per row and/or several in one
"!                cell separated by , ; & / or blanks
"!   column AL    NEW, optional: Package Type ("Loose Pipes", "Boxes",
"!                "Bundles") - entered once per run of lines
"!   rows without Sl.No. (column A) are not line items - e.g. the total
"!   row of the template - and are ignored.
"!   Merged cells (Article No., Articles, Gross/Net Wt.) are supported:
"!   see ZCL_SD_EPACK_RULES, "PACKAGE GROUPS".
CLASS zcl_sd_epack_upload DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "! @parameter iv_mode | ZIF_SD_EPACK=>GC_UPLOAD_MODE-CREATE / -OVERWRITE
    "! @parameter iv_test | Check only, do not save
    METHODS constructor
      IMPORTING iv_mode TYPE char1
                iv_test TYPE abap_bool DEFAULT abap_true.

    "! Read a file from the presentation server and process it
    METHODS upload_file
      IMPORTING iv_path            TYPE string
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

    "! Process file content (entry point independent of SAP GUI)
    METHODS process
      IMPORTING iv_xstring         TYPE xstring
                iv_filename        TYPE string
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

    "! Delete a packing list completely
    CLASS-METHODS delete
      IMPORTING iv_packno          TYPE zsd_packno
                iv_test            TYPE abap_bool DEFAULT abap_true
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

  PRIVATE SECTION.

    TYPES tt_cells TYPE STANDARD TABLE OF string WITH EMPTY KEY.
    TYPES:
      BEGIN OF ty_row,
        row   TYPE i,
        cells TYPE tt_cells,
      END OF ty_row,
      tt_rows TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

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

    DATA mv_mode     TYPE char1.
    DATA mv_test     TYPE abap_bool.
    DATA mt_messages TYPE bapiret2_t.
    DATA ms_header   TYPE zsd_epack_hdr.
    DATA mt_invoices TYPE zif_sd_epack=>tt_vbeln.
    DATA mt_lines    TYPE zif_sd_epack=>tt_upload_line.
    DATA mt_items    TYPE zif_sd_epack=>tt_item.
    DATA mv_bukrs    TYPE bukrs.

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

    METHODS map_header
      IMPORTING it_rows TYPE tt_rows.

    METHODS map_lines
      IMPORTING it_rows TYPE tt_rows.

    METHODS validate.

    METHODS save.

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
      IMPORTING iv_row   TYPE i
                iv_col   TYPE i
                iv_text  TYPE string
                iv_max   TYPE i.

    CLASS-METHODS header_map
      RETURNING VALUE(rt_map) TYPE tt_map.

    CLASS-METHODS column_name
      IMPORTING iv_col         TYPE i
      RETURNING VALUE(rv_name) TYPE string.

    CLASS-METHODS lock
      IMPORTING iv_packno         TYPE zsd_packno
      RETURNING VALUE(rs_message) TYPE bapiret2.

    CLASS-METHODS unlock
      IMPORTING iv_packno TYPE zsd_packno.

ENDCLASS.


CLASS zcl_sd_epack_upload IMPLEMENTATION.

  METHOD constructor.
    mv_mode = iv_mode.
    mv_test = iv_test.
  ENDMETHOD.


  METHOD upload_file.

    DATA lt_solix  TYPE solix_tab.
    DATA lv_length TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING  filename   = iv_path
                 filetype   = 'BIN'
      IMPORTING  filelength = lv_length
      CHANGING   data_tab   = lt_solix
      EXCEPTIONS OTHERS     = 1 ).
    IF sy-subrc <> 0.
      rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '015'
                                                        iv_v1 = iv_path iv_v2 = 'GUI_UPLOAD' ) ) ).
      RETURN.
    ENDIF.

    rt_messages = process( iv_xstring  = cl_bcs_convert=>solix_to_xstring( it_solix = lt_solix
                                                                           iv_size  = lv_length )
                           iv_filename = iv_path ).

  ENDMETHOD.


  METHOD process.

    CLEAR: mt_messages, ms_header, mt_invoices, mt_lines, mt_items, mv_bukrs.

    TRY.
        DATA(lt_rows) = read_worksheet( iv_xstring  = iv_xstring
                                        iv_filename = iv_filename ).
      CATCH zcx_sd_epack INTO DATA(lx_error).
        rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '015'
                                                          iv_v1 = iv_filename
                                                          iv_v2 = lx_error->get_text( ) ) ) ).
        RETURN.
    ENDTRY.

    IF lt_rows IS INITIAL.
      add( iv_no = '036' ).
      rt_messages = mt_messages.
      RETURN.
    ENDIF.

    check_template( lt_rows[ 1 ] ).
    IF zcl_sd_epack_rules=>has_errors( mt_messages ).
      rt_messages = mt_messages.
      RETURN.
    ENDIF.
    DELETE lt_rows INDEX 1.

    map_header( lt_rows ).
    map_lines( lt_rows ).
    ms_header-zsd_filename = iv_filename.

    validate( ).

    IF zcl_sd_epack_rules=>has_errors( mt_messages ).
      add( iv_no = '035' ).
    ELSEIF mv_test = abap_true.
      add( iv_type = 'S' iv_no = '033' iv_v1 = lines( mt_items ) iv_v2 = lines( mt_invoices ) ).
    ELSE.
      save( ).
    ENDIF.

    rt_messages = mt_messages.

  ENDMETHOD.


  METHOD delete.

    SELECT SINGLE zsd_packno, zsd_bukrs FROM zsd_epack_hdr
      WHERE zsd_packno = @iv_packno
      INTO @DATA(ls_hdr).
    IF sy-subrc <> 0.
      rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '007' iv_v1 = iv_packno ) ) ).
      RETURN.
    ENDIF.

    AUTHORITY-CHECK OBJECT zif_sd_epack=>gc_auth_object
      ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-delete
      ID 'BUKRS' FIELD ls_hdr-zsd_bukrs.
    IF sy-subrc <> 0.
      rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '012'
                                                        iv_v1 = 'Company code' iv_v2 = ls_hdr-zsd_bukrs
                                                        iv_v3 = zif_sd_epack=>gc_actvt-delete ) ) ).
      RETURN.
    ENDIF.

    IF iv_test = abap_true.
      rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'S' iv_no = '051' iv_v1 = iv_packno ) ) ).
      RETURN.
    ENDIF.

    DATA(ls_lock) = lock( iv_packno ).
    IF ls_lock IS NOT INITIAL.
      rt_messages = VALUE #( ( ls_lock ) ).
      RETURN.
    ENDIF.

    DELETE FROM zsd_epack_data WHERE zsd_packno = @iv_packno.
    DELETE FROM zsd_epack_inv  WHERE zsd_packno = @iv_packno.
    DELETE FROM zsd_epack_hdr  WHERE zsd_packno = @iv_packno.
    COMMIT WORK AND WAIT.

    unlock( iv_packno ).
    rt_messages = VALUE #( ( zcl_sd_epack_rules=>msg( iv_type = 'S' iv_no = '034' iv_v1 = iv_packno ) ) ).

  ENDMETHOD.


  METHOD add.
    APPEND zcl_sd_epack_rules=>msg( iv_type = iv_type iv_no = iv_no
                                    iv_v1 = iv_v1 iv_v2 = iv_v2
                                    iv_v3 = iv_v3 iv_v4 = iv_v4 ) TO mt_messages.
  ENDMETHOD.


  METHOD read_worksheet.

    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

    " CL_FDT_XL_SPREADSHEET reads .xlsx without SAP GUI / OLE.
    " ⚠ SAP-VERIFY: not a released API; on newer S/4HANA releases
    "   XCO_CP_XLSX is the released alternative (ASSUMPTION A21).
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
    " - catches files in a wrong / shifted layout.
    DATA(lv_a) = to_upper( cell( is_row = is_row iv_col = gc_col-sn ) ).
    DATA(lv_m) = to_upper( cell( is_row = is_row iv_col = gc_col-packno ) ).

    IF lv_a NS 'SL'.
      add( iv_no = '016' iv_v1 = 'A' iv_v2 = cell( is_row = is_row iv_col = gc_col-sn ) ).
    ENDIF.
    IF lv_m NS 'PACKING'.
      add( iv_no = '016' iv_v1 = 'M' iv_v2 = cell( is_row = is_row iv_col = gc_col-packno ) ).
    ENDIF.

  ENDMETHOD.


  METHOD map_header.

    TYPES: BEGIN OF ty_first,
             col   TYPE i,
             row   TYPE i,
             value TYPE string,
           END OF ty_first.
    DATA lt_first TYPE SORTED TABLE OF ty_first WITH UNIQUE KEY col.
    DATA lv_date  TYPE d.
    DATA lv_num   TYPE decfloat34.

    DATA(lt_map) = header_map( ).

    LOOP AT it_rows INTO DATA(ls_row).

      " Invoices: every row may carry one or more numbers
      DATA(lv_inv) = cell( is_row = ls_row iv_col = gc_col-invoice ).
      IF lv_inv IS NOT INITIAL.
        zcl_sd_epack_rules=>split_invoices( EXPORTING iv_text    = lv_inv
                                            IMPORTING et_vbeln   = DATA(lt_vbeln)
                                                      et_invalid = DATA(lt_invalid) ).
        INSERT LINES OF lt_vbeln INTO TABLE mt_invoices.
        LOOP AT lt_invalid INTO DATA(lv_invalid).
          add( iv_no = '018' iv_v1 = ls_row-row iv_v2 = lv_invalid iv_v3 = 'L' ).
        ENDLOOP.
      ENDIF.

      LOOP AT lt_map INTO DATA(ls_map).
        DATA(lv_value) = cell( is_row = ls_row iv_col = ls_map-col ).
        IF lv_value IS INITIAL.
          CONTINUE.
        ENDIF.

        " Same value in several rows is fine, a different one is not
        READ TABLE lt_first INTO DATA(ls_first) WITH TABLE KEY col = ls_map-col.
        IF sy-subrc = 0.
          IF ls_first-value <> lv_value.
            add( iv_no = '020' iv_v1 = ls_row-row iv_v2 = ls_map-label iv_v3 = ls_first-row ).
          ENDIF.
          CONTINUE.
        ENDIF.
        INSERT VALUE #( col = ls_map-col row = ls_row-row value = lv_value ) INTO TABLE lt_first.

        ASSIGN COMPONENT ls_map-field OF STRUCTURE ms_header TO FIELD-SYMBOL(<lv_field>).
        CHECK sy-subrc = 0.

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
            " Keep Alt+Enter line breaks of multi-line cells (kind T)
            check_length( iv_row = ls_row-row iv_col = ls_map-col iv_text = lv_value iv_max = ls_map-maxlen ).
            <lv_field> = lv_value.
        ENDCASE.
      ENDLOOP.

    ENDLOOP.

    ms_header-zsd_packno = |{ ms_header-zsd_packno ALPHA = IN }|.

    IF ms_header-zsd_packno IS INITIAL.
      add( iv_no = '017' iv_v1 = 2 iv_v2 = 'Packing list no' ).
    ENDIF.
    IF ms_header-zsd_packdt IS INITIAL.
      add( iv_no = '017' iv_v1 = 2 iv_v2 = 'Packing list date' ).
    ENDIF.
    IF mt_invoices IS INITIAL.
      add( iv_no = '017' iv_v1 = 2 iv_v2 = 'SAP Invoice Number' ).
    ENDIF.

    " Notify parties must be filled from 1 upwards - the form prints
    " "1ST NOTIFY PARTY" ... (ASSUMPTION A10)
    DATA lv_gap TYPE i.
    DO zif_sd_epack=>gc_max_np TIMES.
      ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE ms_header TO FIELD-SYMBOL(<lv_np>).
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

    DATA lv_ignored TYPE i.

    LOOP AT it_rows INTO DATA(ls_row).

      " A line item needs a Sl.No.; other rows (total row of the
      " template, header-only rows) are skipped
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
      ls_line-art    = number( is_row = ls_row iv_col = gc_col-art iv_integer = abap_true ).
      ls_line-qty    = round( val = number( is_row = ls_row iv_col = gc_col-qty ) dec = 3 ).
      ls_line-gwt    = round( val = number( is_row = ls_row iv_col = gc_col-gwt ) dec = 3 ).
      ls_line-nwt    = round( val = number( is_row = ls_row iv_col = gc_col-nwt ) dec = 3 ).

      " Empty cell = merged cell (part of the group above), see A06
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

      APPEND ls_line TO mt_lines.
    ENDLOOP.

    IF lv_ignored > 0.
      add( iv_type = 'I' iv_no = '050' iv_v1 = lv_ignored ).
    ENDIF.

    IF mt_lines IS INITIAL.
      add( iv_no = '036' ).
      RETURN.
    ENDIF.

    zcl_sd_epack_rules=>build_items( EXPORTING iv_packno   = ms_header-zsd_packno
                                               it_lines    = mt_lines
                                     IMPORTING et_items    = mt_items
                                     CHANGING  ct_messages = mt_messages ).

  ENDMETHOD.


  METHOD validate.

    " --- Business rules on the line items --------------------------------
    APPEND LINES OF zcl_sd_epack_rules=>check_weights( mt_items )    TO mt_messages.
    APPEND LINES OF zcl_sd_epack_rules=>check_quantities( mt_items ) TO mt_messages.

    IF mt_invoices IS INITIAL OR ms_header-zsd_packno IS INITIAL.
      RETURN.
    ENDIF.

    " --- Invoices ---------------------------------------------------------
    DATA lr_fkart TYPE RANGE OF fkart.

    SELECT vbeln, fkart, fksto, sfakn, bukrs FROM vbrk
      FOR ALL ENTRIES IN @mt_invoices
      WHERE vbeln = @mt_invoices-table_line
      INTO TABLE @DATA(lt_vbrk).
    SORT lt_vbrk BY vbeln.

    SELECT sign, opti AS option, low, high FROM tvarvc
      WHERE name = @zif_sd_epack=>gc_tvarv_fkart
        AND type = 'S'
      INTO CORRESPONDING FIELDS OF TABLE @lr_fkart.
    DELETE lr_fkart WHERE low IS INITIAL AND high IS INITIAL.

    LOOP AT mt_invoices INTO DATA(lv_vbeln).
      DATA(lv_ext) = |{ lv_vbeln ALPHA = OUT }|.
      READ TABLE lt_vbrk INTO DATA(ls_vbrk) WITH KEY vbeln = lv_vbeln BINARY SEARCH.
      IF sy-subrc <> 0.
        add( iv_no = '008' iv_v1 = lv_ext ).
        CONTINUE.
      ENDIF.
      IF ls_vbrk-fksto = abap_true OR ls_vbrk-sfakn IS NOT INITIAL.
        add( iv_no = '009' iv_v1 = lv_ext ).
      ENDIF.
      IF lr_fkart IS NOT INITIAL AND ls_vbrk-fkart NOT IN lr_fkart.
        add( iv_no = '010' iv_v1 = lv_ext iv_v2 = ls_vbrk-fkart ).
      ENDIF.
      IF mv_bukrs IS INITIAL.
        mv_bukrs = ls_vbrk-bukrs.
      ELSEIF ls_vbrk-bukrs <> mv_bukrs.
        add( iv_no = '011' iv_v1 = 'Company code' iv_v2 = mv_bukrs iv_v3 = ls_vbrk-bukrs ).
      ENDIF.
    ENDLOOP.
    ms_header-zsd_bukrs = mv_bukrs.

    " One invoice belongs to one packing list (ASSUMPTION A01)
    SELECT zsd_packno, vbeln FROM zsd_epack_inv
      FOR ALL ENTRIES IN @mt_invoices
      WHERE vbeln = @mt_invoices-table_line
      INTO TABLE @DATA(lt_link).
    LOOP AT lt_link INTO DATA(ls_link) WHERE zsd_packno <> ms_header-zsd_packno.
      add( iv_no = '026' iv_v1 = |{ ls_link-vbeln ALPHA = OUT }| iv_v2 = ls_link-zsd_packno ).
    ENDLOOP.

    " Part No. should be a billed material (warning only, ASSUMPTION A12)
    SELECT DISTINCT matnr FROM vbrp
      FOR ALL ENTRIES IN @mt_invoices
      WHERE vbeln = @mt_invoices-table_line
      INTO TABLE @DATA(lt_matnr).
    LOOP AT mt_items INTO DATA(ls_item).
      DATA lv_matnr TYPE matnr.
      CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
        EXPORTING
          input        = ls_item-zsd_partno
        IMPORTING
          output       = lv_matnr
        EXCEPTIONS
          length_error = 1
          OTHERS       = 2.
      IF sy-subrc <> 0 OR NOT line_exists( lt_matnr[ matnr = lv_matnr ] ).
        add( iv_type = 'W' iv_no = '030' iv_v1 = |{ CONV i( ls_item-zsd_sn ) }| iv_v2 = ls_item-zsd_partno ).
      ENDIF.
    ENDLOOP.

    " --- Packing list existence vs. mode ----------------------------------
    SELECT SINGLE zsd_revno FROM zsd_epack_hdr
      WHERE zsd_packno = @ms_header-zsd_packno
      INTO @DATA(lv_revno).
    DATA(lv_exists) = xsdbool( sy-subrc = 0 ).

    CASE mv_mode.
      WHEN zif_sd_epack=>gc_upload_mode-create.
        IF lv_exists = abap_true.
          add( iv_no = '024' iv_v1 = ms_header-zsd_packno ).
        ENDIF.
      WHEN zif_sd_epack=>gc_upload_mode-overwrite.
        IF lv_exists = abap_false.
          add( iv_no = '025' iv_v1 = ms_header-zsd_packno ).
        ELSE.
          ms_header-zsd_revno = lv_revno + 1.
        ENDIF.
    ENDCASE.

    " --- Authorization (ASSUMPTION A24) ------------------------------------
    DATA(lv_actvt) = COND activ_auth( WHEN mv_mode = zif_sd_epack=>gc_upload_mode-create
                                      THEN zif_sd_epack=>gc_actvt-create
                                      ELSE zif_sd_epack=>gc_actvt-change ).
    AUTHORITY-CHECK OBJECT zif_sd_epack=>gc_auth_object
      ID 'ACTVT' FIELD lv_actvt
      ID 'BUKRS' FIELD mv_bukrs.
    IF sy-subrc <> 0.
      add( iv_no = '012' iv_v1 = 'Company code' iv_v2 = mv_bukrs iv_v3 = lv_actvt ).
    ENDIF.

  ENDMETHOD.


  METHOD save.

    DATA(ls_lock) = lock( ms_header-zsd_packno ).
    IF ls_lock IS NOT INITIAL.
      APPEND ls_lock TO mt_messages.
      RETURN.
    ENDIF.

    " Audit fields - keep the creation data on overwrite (ASSUMPTION A22)
    SELECT SINGLE zsd_ernam, zsd_erdat, zsd_erzet FROM zsd_epack_hdr
      WHERE zsd_packno = @ms_header-zsd_packno
      INTO (@ms_header-zsd_ernam, @ms_header-zsd_erdat, @ms_header-zsd_erzet).
    IF sy-subrc <> 0.
      ms_header-zsd_ernam = sy-uname.
      ms_header-zsd_erdat = sy-datum.
      ms_header-zsd_erzet = sy-uzeit.
    ELSE.
      ms_header-zsd_aenam = sy-uname.
      ms_header-zsd_aedat = sy-datum.
      ms_header-zsd_aezet = sy-uzeit.
    ENDIF.

    DATA lt_inv_db TYPE STANDARD TABLE OF zsd_epack_inv WITH EMPTY KEY.
    lt_inv_db = VALUE #( FOR lv_vbeln IN mt_invoices
                         ( zsd_packno = ms_header-zsd_packno vbeln = lv_vbeln ) ).

    DELETE FROM zsd_epack_data WHERE zsd_packno = @ms_header-zsd_packno.
    DELETE FROM zsd_epack_inv  WHERE zsd_packno = @ms_header-zsd_packno.

    MODIFY zsd_epack_hdr FROM @ms_header.
    DATA(lv_ok) = xsdbool( sy-subrc = 0 ).

    INSERT zsd_epack_inv FROM TABLE @lt_inv_db.
    IF sy-subrc <> 0.
      lv_ok = abap_false.
    ENDIF.

    INSERT zsd_epack_data FROM TABLE @mt_items.
    IF sy-subrc <> 0.
      lv_ok = abap_false.
    ENDIF.

    IF lv_ok = abap_true.
      COMMIT WORK AND WAIT.
      add( iv_type = 'S' iv_no = '032' iv_v1 = ms_header-zsd_packno iv_v2 = ms_header-zsd_revno
                                       iv_v3 = lines( mt_items ) iv_v4 = lines( mt_invoices ) ).
    ELSE.
      ROLLBACK WORK.
      add( iv_no = '052' iv_v1 = ms_header-zsd_packno ).
    ENDIF.

    unlock( ms_header-zsd_packno ).

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
    " Lengths of multi-line fields are wider than in the FS - the sample
    " outputs do not fit into the FS lengths (ASSUMPTION A19).
    rt_map = VALUE #(
      ( col = 13 field = 'ZSD_PACKNO'    kind = 'C' maxlen = 10   label = `Packing list no` )
      ( col = 14 field = 'ZSD_PACKDT'    kind = 'D'               label = `Packing list date` )
      ( col = 15 field = 'ZSD_EXPREF'    kind = 'C' maxlen = 20   label = `Exporter Ref` )
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


  METHOD lock.

    CALL FUNCTION 'ENQUEUE_EZSD_EPACK'
      EXPORTING
        mode_zsd_epack_hdr = 'E'
        mandt              = sy-mandt
        zsd_packno         = iv_packno
      EXCEPTIONS
        foreign_lock       = 1
        system_failure     = 2
        OTHERS             = 3.
    IF sy-subrc <> 0.
      rs_message = zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '014' iv_v1 = iv_packno iv_v2 = sy-msgv1 ).
    ENDIF.

  ENDMETHOD.


  METHOD unlock.

    CALL FUNCTION 'DEQUEUE_EZSD_EPACK'
      EXPORTING
        mode_zsd_epack_hdr = 'E'
        mandt              = sy-mandt
        zsd_packno         = iv_packno.

  ENDMETHOD.

ENDCLASS.
