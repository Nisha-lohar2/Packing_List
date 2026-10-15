*&---------------------------------------------------------------------*
*& Report ZSD_EPACK_UPLOAD                    (T-code ZSD_EPACK_UPL *)
*&---------------------------------------------------------------------*
*& WRICEF 102-B  Export consolidated packing list - REPORT 2
*&
*& 1. Upload of the Excel file written by report 1 ZSD_EPACK_EXTRACT and
*&    completed by the user. Path and format (.xlsx) are checked.
*& 2. The first worksheet is read into an internal table typed like
*&    ZSD_EPACK_DATA. Columns are mapped by POSITION and the heading of
*&    every column is checked against the layout (FORM epack_columns).
*&    The Excel row number is kept for every row.
*& 3. Validations (v3/Validation_Mapping.md): mandatory / key fields,
*&    data types and lengths, duplicate keys, sales order exists, Part No.
*&    is an item of the sales order, uploaded quantity <= delivered
*&    quantity (sheet "Validation"), gross weight >= net weight (FS),
*&    Smart Form format vs. notifier parties (FS), packing list already
*&    saved. All errors are collected with row and column.
*& 4. Save to ZSD_EPACK_DATA - all or nothing (default) or per packing
*&    list, lock per packing list, COMMIT / ROLLBACK. A packing list that
*&    is already saved is only replaced when "Overwrite" is ticked, so a
*&    rerun with the same file never creates duplicates.
*& 5. Smart Form per saved packing list (with / without notifier party).
*& Test run: steps 1 - 3 only.
*&
*& * T-code of report 2 to be confirmed (open point Q02).
*&---------------------------------------------------------------------*
REPORT zsd_epack_upload.

INCLUDE zsd_epack_common.

*----------------------------------------------------------------------*
* Constants
*----------------------------------------------------------------------*
CONSTANTS:
  " Smart Forms - names to be confirmed (open point Q09)
  gc_form_np   TYPE tdsfname VALUE 'ZSD_EPACK_SF_NP',  " with notifier party
  gc_form_wonp TYPE tdsfname VALUE 'ZSD_EPACK_SF',     " without notifier party
  gc_gewei     TYPE gewei VALUE 'KG',                  " "Gross / Net Wt. KGS"
  gc_meins     TYPE meins VALUE 'ST',                  " "Quantity Pcs" if the order unit is not found
  gc_msg_qty   TYPE string VALUE `Uploaded Quantity should not be greater than Delivered Quantity`,
  gc_msg_wt    TYPE string VALUE `Invalid Weight: Gross Weight cannot be less than Net Weight.`.

*----------------------------------------------------------------------*
* Types
*----------------------------------------------------------------------*
TYPES:
  " Uploaded row: Excel row number + data typed like the Z table
  BEGIN OF ty_upl,
    row   TYPE i,
    error TYPE abap_bool,
    matnr TYPE matnr,                         " Part No. in internal format
    data  TYPE zsd_epack_data,
  END OF ty_upl,
  ty_upl_tab TYPE STANDARD TABLE OF ty_upl WITH EMPTY KEY,

  ty_db_tab TYPE STANDARD TABLE OF zsd_epack_data WITH EMPTY KEY,

  " Packing list status
  BEGIN OF ty_pl,
    packno    TYPE zsd_epack_data-zsd_packno,
    first_row TYPE i,
    error     TYPE abap_bool,
    exists    TYPE abap_bool,
    locked    TYPE abap_bool,
    saved     TYPE abap_bool,
    printed   TYPE abap_bool,
  END OF ty_pl,
  ty_pl_tab TYPE SORTED TABLE OF ty_pl WITH UNIQUE KEY packno,

  " Sales order of an upload row (a row may name several, separated by ",")
  BEGIN OF ty_row_so,
    idx   TYPE i,                             " index in GT_UPL
    vbeln TYPE vbak-vbeln,
  END OF ty_row_so,
  ty_row_so_tab TYPE SORTED TABLE OF ty_row_so WITH UNIQUE KEY idx vbeln,

  " Quantity per sales order and material
  BEGIN OF ty_so_mat,
    vbeln TYPE vbak-vbeln,
    matnr TYPE matnr,
    qty   TYPE p LENGTH 16 DECIMALS 3,
  END OF ty_so_mat,
  ty_so_mat_tab TYPE HASHED TABLE OF ty_so_mat WITH UNIQUE KEY vbeln matnr,

  " Order item unit per sales order and material
  BEGIN OF ty_so_unit,
    vbeln TYPE vbak-vbeln,
    matnr TYPE matnr,
    vrkme TYPE vbap-vrkme,
  END OF ty_so_unit,
  ty_so_unit_tab TYPE HASHED TABLE OF ty_so_unit WITH UNIQUE KEY vbeln matnr,

  " Quantity check group: sales order(s) of the row + material
  BEGIN OF ty_group,
    key      TYPE string,
    uploaded TYPE p LENGTH 16 DECIMALS 3,
    delivered TYPE p LENGTH 16 DECIMALS 3,
  END OF ty_group,
  ty_group_tab TYPE HASHED TABLE OF ty_group WITH UNIQUE KEY key.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
DATA gt_cols TYPE ty_col_tab.           " Excel layout
DATA gt_upl  TYPE ty_upl_tab.           " uploaded rows
DATA gt_pl   TYPE ty_pl_tab.            " packing lists of the file
DATA gt_old  TYPE ty_db_tab.            " keys already in ZSD_EPACK_DATA
DATA gt_msg  TYPE ty_msg_tab.           " messages
DATA gv_stop TYPE abap_bool.            " file cannot be processed

*----------------------------------------------------------------------*
* Selection screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_file TYPE rlgrap-filename LOWER CASE OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

" FS: two Smart Form formats; with notifier party when NP are maintained
SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-b02.
PARAMETERS: p_wonp RADIOBUTTON GROUP fmt DEFAULT 'X',
            p_wnp  RADIOBUTTON GROUP fmt.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-b03.
PARAMETERS: p_all  RADIOBUTTON GROUP sav DEFAULT 'X',   " all or nothing
            p_part RADIOBUTTON GROUP sav,               " per packing list
            p_ovwr AS CHECKBOX,                         " replace saved packing list
            p_test AS CHECKBOX DEFAULT 'X'.             " validate only
SELECTION-SCREEN END OF BLOCK b3.

*----------------------------------------------------------------------*
INITIALIZATION.
*----------------------------------------------------------------------*
  PERFORM epack_columns CHANGING gt_cols.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_file CHANGING p_file.

AT SELECTION-SCREEN ON p_file.
  IF sy-ucomm = 'ONLI' OR sy-ucomm = 'PRIN'.
    PERFORM check_file.
  ENDIF.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  IF sy-batch = abap_true.
    MESSAGE 'The file is read from the PC - run the report in dialog' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  PERFORM read_file.                       " file -> GT_UPL, headings checked
  IF gv_stop = abap_false.
    PERFORM check_keys.                    " mandatory, duplicates
    PERFORM check_orders.                  " sales order, material, quantity
    PERFORM check_weights.                 " GW >= NW
    PERFORM check_packing_lists.           " format / NP, existing, header
    PERFORM decide.                        " what can be saved
    IF p_test = abap_false.
      PERFORM save_data.
      PERFORM print_forms.
    ENDIF.
  ENDIF.
  PERFORM show_result.

*&---------------------------------------------------------------------*
*& File path and format
*&---------------------------------------------------------------------*
FORM check_file.

  DATA lv_file  TYPE string.
  DATA lv_exist TYPE abap_bool.

  lv_file = p_file.
  IF to_upper( substring_after( val = lv_file sub = `.` occ = -1 ) ) <> `XLSX`.
    MESSAGE 'Only Excel files .xlsx can be uploaded (file of report ZSD_EPACK_EXTRACT)' TYPE 'E'.
  ENDIF.

  cl_gui_frontend_services=>file_exist(
    EXPORTING
      file   = lv_file
    RECEIVING
      result = lv_exist
    EXCEPTIONS
      OTHERS = 1 ).
  IF sy-subrc <> 0 OR lv_exist = abap_false.
    MESSAGE |File { lv_file } does not exist| TYPE 'E'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel -> GT_UPL. Row 1 = headings (checked per position), data from
*& row 2. Every cell is converted to the type of its Z-table field.
*&---------------------------------------------------------------------*
FORM read_file.

  DATA lr_data   TYPE REF TO data.
  DATA ls_upl    TYPE ty_upl.
  DATA lv_text   TYPE string.
  DATA lv_found  TYPE string.
  DATA lv_expect TYPE string.
  DATA lv_alt    TYPE string.
  DATA lv_error  TYPE string.
  DATA lv_letter TYPE string.
  DATA lv_row    TYPE i.
  DATA lv_empty  TYPE abap_bool.
  DATA lv_pack2  TYPE zsd_epack_data-zsd_packno.
  DATA lv_pack2t TYPE string.
  DATA lv_next   TYPE i.

  FIELD-SYMBOLS <lt_excel> TYPE STANDARD TABLE.
  FIELD-SYMBOLS <ls_excel> TYPE any.
  FIELD-SYMBOLS <lv_cell>  TYPE any.
  FIELD-SYMBOLS <lv_field> TYPE any.

  PERFORM read_xlsx CHANGING lr_data.
  IF gv_stop = abap_true.
    RETURN.
  ENDIF.
  ASSIGN lr_data->* TO <lt_excel>.
  IF <lt_excel> IS NOT ASSIGNED OR lines( <lt_excel> ) < 2.
    PERFORM add_msg USING 0 space 0 'E' 'The file has no data below the heading row'.
    gv_stop = abap_true.
    RETURN.
  ENDIF.

  " ---- Row 1: headings, position by position ------------------------------
  READ TABLE <lt_excel> ASSIGNING <ls_excel> INDEX 1.
  LOOP AT gt_cols INTO DATA(ls_col).
    CLEAR lv_text.
    ASSIGN COMPONENT ls_col-pos OF STRUCTURE <ls_excel> TO <lv_cell>.
    IF sy-subrc = 0.
      lv_text = <lv_cell>.
    ENDIF.
    PERFORM normalize USING lv_text CHANGING lv_found.
    PERFORM normalize USING ls_col-header CHANGING lv_expect.
    " Key columns also accepted without the words "KEY FIELD"
    lv_alt = replace( val = lv_expect sub = `KEYFIELD` with = `` ).
    IF lv_found <> lv_expect AND lv_found <> lv_alt.
      lv_text = |Heading "{ lv_text }" found, "{ ls_col-header }" expected - use the file of report ZSD_EPACK_EXTRACT|.
      PERFORM add_msg USING 1 space ls_col-pos 'E' lv_text.
      gv_stop = abap_true.
    ENDIF.
  ENDLOOP.
  " Columns after the layout must be empty
  lv_next = lines( gt_cols ) + 1.
  ASSIGN COMPONENT lv_next OF STRUCTURE <ls_excel> TO <lv_cell>.
  IF sy-subrc = 0 AND <lv_cell> IS NOT INITIAL.
    PERFORM col_letter USING lv_next CHANGING lv_letter.
    lv_text = |Column { lv_letter } and following are not part of the layout - ignored|.
    PERFORM add_msg USING 1 space 0 'W' lv_text.
  ENDIF.
  IF gv_stop = abap_true.
    RETURN.
  ENDIF.

  " ---- Data rows ---------------------------------------------------------
  LOOP AT <lt_excel> ASSIGNING <ls_excel> FROM 2.
    lv_row = sy-tabix.                                  " Excel row number
    CLEAR ls_upl.
    ls_upl-row = lv_row.

    " Empty row -> skipped
    lv_empty = abap_true.
    LOOP AT gt_cols INTO ls_col.
      ASSIGN COMPONENT ls_col-pos OF STRUCTURE <ls_excel> TO <lv_cell>.
      IF sy-subrc = 0 AND condense( CONV string( <lv_cell> ) ) IS NOT INITIAL.
        lv_empty = abap_false.
        EXIT.
      ENDIF.
    ENDLOOP.
    IF lv_empty = abap_true.
      CONTINUE.
    ENDIF.

    CLEAR lv_pack2t.
    LOOP AT gt_cols INTO ls_col.
      CLEAR lv_text.
      ASSIGN COMPONENT ls_col-pos OF STRUCTURE <ls_excel> TO <lv_cell>.
      IF sy-subrc = 0.
        lv_text = <lv_cell>.
      ENDIF.
      IF ls_col-source = gc_src_copy.
        lv_pack2t = lv_text.                            " checked below
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_col-field OF STRUCTURE ls_upl-data TO <lv_field>.
      PERFORM text_to_field USING lv_text CHANGING <lv_field> lv_error.
      IF lv_error IS NOT INITIAL.
        ls_upl-error = abap_true.
        PERFORM add_msg USING lv_row space ls_col-pos 'E' lv_error.
      ENDIF.
    ENDLOOP.

    " Column P (2nd "Packing list no") must repeat column A
    PERFORM text_to_field USING lv_pack2t CHANGING lv_pack2 lv_error.
    IF lv_pack2 IS NOT INITIAL AND lv_pack2 <> ls_upl-data-zsd_packno.
      ls_upl-error = abap_true.
      READ TABLE gt_cols INTO ls_col WITH KEY source = gc_src_copy.
      lv_text = |"{ lv_pack2 }" differs from the packing list no. of the key column "{ ls_upl-data-zsd_packno }"|.
      PERFORM add_msg USING lv_row ls_upl-data-zsd_packno ls_col-pos 'E' lv_text.
    ENDIF.

    " Sales order(s): internal format, sorted, separated by ","
    PERFORM normalize_orders CHANGING ls_upl.

    ls_upl-data-zsd_gewei = gc_gewei.
    APPEND ls_upl TO gt_upl.
  ENDLOOP.

  IF gt_upl IS INITIAL.
    PERFORM add_msg USING 0 space 0 'E' 'The file has no data rows'.
    gv_stop = abap_true.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Front-end file -> first worksheet as internal table (one STRING
*& component per Excel column, row 1 = Excel row 1)
*&---------------------------------------------------------------------*
FORM read_xlsx CHANGING cr_data TYPE REF TO data.

  DATA lt_bin     TYPE solix_tab.
  DATA lv_size    TYPE i.
  DATA lv_xstring TYPE xstring.
  DATA lv_file    TYPE string.
  DATA lv_text    TYPE string.
  DATA lo_excel   TYPE REF TO cl_fdt_xl_spreadsheet.
  DATA lt_sheets  TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.

  lv_file = p_file.
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename   = lv_file
      filetype   = 'BIN'
    IMPORTING
      filelength = lv_size
    CHANGING
      data_tab   = lt_bin
    EXCEPTIONS
      OTHERS     = 1 ).
  IF sy-subrc <> 0 OR lv_size = 0.
    lv_text = |File { lv_file } could not be read|.
    PERFORM add_msg USING 0 space 0 'E' lv_text.
    gv_stop = abap_true.
    RETURN.
  ENDIF.
  lv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_bin iv_size = lv_size ).

  TRY.
      CREATE OBJECT lo_excel
        EXPORTING
          document_name = lv_file
          xdocument     = lv_xstring.
      lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
        IMPORTING worksheet_names = lt_sheets ).
      IF lt_sheets IS INITIAL.
        PERFORM add_msg USING 0 space 0 'E' 'The Excel file has no worksheet'.
        gv_stop = abap_true.
        RETURN.
      ENDIF.
      cr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lt_sheets[ 1 ] ).
    CATCH cx_fdt_excel_core.
      PERFORM add_msg USING 0 space 0 'E' 'The file is not a valid Excel (.xlsx) file'.
      gv_stop = abap_true.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& "SAP Sales Order No" -> internal format. Several orders may be given
*& separated by "," ";" or blanks (field length 255, A05).
*&---------------------------------------------------------------------*
FORM normalize_orders CHANGING cs_upl TYPE ty_upl.

  DATA lt_parts TYPE ty_string_tab.
  DATA lt_vbeln TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.
  DATA lv_vbeln TYPE vbak-vbeln.
  DATA lv_text  TYPE string.
  DATA lv_pos   TYPE i.

  IF cs_upl-data-vbak_vbeln IS INITIAL.
    RETURN.
  ENDIF.
  READ TABLE gt_cols INTO DATA(ls_col) WITH KEY field = 'VBAK_VBELN'.
  lv_pos = ls_col-pos.

  lv_text = cs_upl-data-vbak_vbeln.
  REPLACE ALL OCCURRENCES OF REGEX `[,;\s]+` IN lv_text WITH `,`.
  SPLIT lv_text AT `,` INTO TABLE lt_parts.
  DELETE lt_parts WHERE table_line IS INITIAL.

  LOOP AT lt_parts INTO DATA(lv_part).
    IF strlen( lv_part ) > 10 OR lv_part CN '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'.
      cs_upl-error = abap_true.
      lv_text = |"{ lv_part }" is not a sales order number|.
      PERFORM add_msg USING cs_upl-row cs_upl-data-zsd_packno lv_pos 'E' lv_text.
      CONTINUE.
    ENDIF.
    lv_vbeln = to_upper( lv_part ).
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = lv_vbeln
      IMPORTING
        output = lv_vbeln.
    INSERT lv_vbeln INTO TABLE lt_vbeln.
  ENDLOOP.

  CLEAR lv_text.
  LOOP AT lt_vbeln INTO lv_vbeln.
    lv_text = COND #( WHEN lv_text IS INITIAL THEN |{ lv_vbeln }| ELSE |{ lv_text },{ lv_vbeln }| ).
  ENDLOOP.
  IF strlen( lv_text ) > 255.
    cs_upl-error = abap_true.
    PERFORM add_msg USING cs_upl-row cs_upl-data-zsd_packno lv_pos 'E' 'Too many sales orders for 255 characters'.
    RETURN.
  ENDIF.
  cs_upl-data-vbak_vbeln = lv_text.

ENDFORM.

*&---------------------------------------------------------------------*
*& Key fields mandatory, no duplicate key in the file
*&---------------------------------------------------------------------*
FORM check_keys.

  TYPES: BEGIN OF ty_key,
           packno TYPE zsd_epack_data-zsd_packno,
           vbeln  TYPE zsd_epack_data-vbak_vbeln,
           sn     TYPE zsd_epack_data-zsd_sn,
           idx    TYPE i,
         END OF ty_key.

  DATA lt_key  TYPE STANDARD TABLE OF ty_key WITH EMPTY KEY.
  DATA ls_prev TYPE ty_key.
  DATA lv_text TYPE string.

  LOOP AT gt_upl ASSIGNING FIELD-SYMBOL(<ls_upl>).
    DATA(lv_idx) = sy-tabix.
    IF <ls_upl>-data-zsd_packno IS INITIAL.
      PERFORM row_error USING lv_idx 'ZSD_PACKNO' 'Packing list no is mandatory (key field)'.
    ENDIF.
    IF <ls_upl>-data-vbak_vbeln IS INITIAL.
      PERFORM row_error USING lv_idx 'VBAK_VBELN' 'SAP Sales Order No is mandatory (key field)'.
    ENDIF.
    IF <ls_upl>-data-zsd_sn IS INITIAL.
      PERFORM row_error USING lv_idx 'ZSD_SN' 'Sl.No. is mandatory (line key, A03)'.
    ENDIF.
    APPEND VALUE #( packno = <ls_upl>-data-zsd_packno vbeln = <ls_upl>-data-vbak_vbeln
                    sn = <ls_upl>-data-zsd_sn idx = lv_idx ) TO lt_key.
  ENDLOOP.

  SORT lt_key BY packno vbeln sn idx.
  LOOP AT lt_key INTO DATA(ls_key) WHERE packno IS NOT INITIAL AND sn IS NOT INITIAL.
    IF ls_key-packno = ls_prev-packno AND ls_key-vbeln = ls_prev-vbeln AND ls_key-sn = ls_prev-sn.
      lv_text = |Packing list no / Sales order / Sl.No. { ls_key-sn ALPHA = OUT } also in Excel row { gt_upl[ ls_prev-idx ]-row }|.
      PERFORM row_error USING ls_key-idx 'ZSD_SN' lv_text.
    ELSE.
      ls_prev = ls_key.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Sheet "Validation": sales order exists; Part No. is an item of the
*& sales order (VBAP); uploaded quantity <= delivered quantity (LIPS
*& LFIMG by VGBEL = sales order). Quantities are added up per sales
*& order(s) + material over all rows of the file (A06).
*&---------------------------------------------------------------------*
FORM check_orders.

  TYPES: BEGIN OF ty_lips,
           vbeln TYPE lips-vbeln,
           posnr TYPE lips-posnr,
           vgbel TYPE lips-vgbel,
           matnr TYPE lips-matnr,
           lfimg TYPE lips-lfimg,
         END OF ty_lips.

  DATA lt_row_so  TYPE ty_row_so_tab.
  DATA lt_so      TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.
  DATA lt_exist   TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.
  DATA lt_unit    TYPE ty_so_unit_tab.
  DATA lt_deliv   TYPE ty_so_mat_tab.
  DATA lt_group   TYPE ty_group_tab.
  DATA lt_lips    TYPE STANDARD TABLE OF ty_lips WITH EMPTY KEY.
  DATA lt_rowkey  TYPE STANDARD TABLE OF string WITH EMPTY KEY.   " group key per GT_UPL index
  DATA lt_parts   TYPE ty_string_tab.
  DATA ls_group   TYPE ty_group.
  DATA ls_so_mat  TYPE ty_so_mat.
  DATA lv_vbeln   TYPE vbak-vbeln.
  DATA lv_text    TYPE string.
  DATA lv_found   TYPE abap_bool.
  DATA lv_partno  TYPE c LENGTH 40.

  " ---- Sales orders of every row ------------------------------------------
  LOOP AT gt_upl ASSIGNING FIELD-SYMBOL(<ls_upl>).
    DATA(lv_idx) = sy-tabix.
    IF <ls_upl>-data-vbak_vbeln IS INITIAL.
      CONTINUE.
    ENDIF.
    lv_text = <ls_upl>-data-vbak_vbeln.
    SPLIT lv_text AT `,` INTO TABLE lt_parts.
    LOOP AT lt_parts INTO DATA(lv_part).
      lv_vbeln = lv_part.
      INSERT VALUE #( idx = lv_idx vbeln = lv_vbeln ) INTO TABLE lt_row_so.
      INSERT lv_vbeln INTO TABLE lt_so.
    ENDLOOP.
  ENDLOOP.
  IF lt_so IS INITIAL.
    RETURN.
  ENDIF.

  " ---- Bulk reads ------------------------------------------------------------
  SELECT vbeln
    FROM vbak
    FOR ALL ENTRIES IN @lt_so
    WHERE vbeln = @lt_so-table_line
    INTO TABLE @lt_exist.

  " Sales order items: material and sales unit (sheet: VBAP MATNR, KWMENG)
  SELECT vbeln, matnr, vrkme
    FROM vbap
    FOR ALL ENTRIES IN @lt_so
    WHERE vbeln = @lt_so-table_line
    INTO TABLE @DATA(lt_vbap).
  LOOP AT lt_vbap INTO DATA(ls_vbap) WHERE matnr IS NOT INITIAL.
    INSERT VALUE #( vbeln = ls_vbap-vbeln matnr = ls_vbap-matnr vrkme = ls_vbap-vrkme ) INTO TABLE lt_unit.
  ENDLOOP.

  " Delivery items of the sales orders (sheet: LIPS VGBEL -> MATNR, LFIMG)
  SELECT vbeln, posnr, vgbel, matnr, lfimg
    FROM lips
    FOR ALL ENTRIES IN @lt_so
    WHERE vgbel = @lt_so-table_line
    INTO TABLE @lt_lips.
  LOOP AT lt_lips INTO DATA(ls_lips).
    ls_so_mat = VALUE #( vbeln = ls_lips-vgbel matnr = ls_lips-matnr qty = ls_lips-lfimg ).
    COLLECT ls_so_mat INTO lt_deliv.
  ENDLOOP.

  " ---- Per row --------------------------------------------------------------
  CLEAR lt_rowkey.
  LOOP AT gt_upl ASSIGNING <ls_upl>.
    lv_idx = sy-tabix.
    APPEND INITIAL LINE TO lt_rowkey ASSIGNING FIELD-SYMBOL(<lv_key>).
    IF <ls_upl>-data-vbak_vbeln IS INITIAL.
      CONTINUE.
    ENDIF.

    " Sales order exists
    LOOP AT lt_row_so INTO DATA(ls_row_so) WHERE idx = lv_idx.
      IF NOT line_exists( lt_exist[ table_line = ls_row_so-vbeln ] ).
        lv_text = |Sales order { ls_row_so-vbeln ALPHA = OUT } does not exist|.
        PERFORM row_error USING lv_idx 'VBAK_VBELN' lv_text.
      ENDIF.
    ENDLOOP.

    " Part No. -> material number
    IF <ls_upl>-data-zsd_partno IS NOT INITIAL.
      lv_partno = <ls_upl>-data-zsd_partno.
      CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
        EXPORTING
          input        = lv_partno
        IMPORTING
          output       = <ls_upl>-matnr
        EXCEPTIONS
          length_error = 1
          OTHERS       = 2.
      IF sy-subrc <> 0.
        PERFORM row_error USING lv_idx 'ZSD_PARTNO' 'Part No. is not a valid material number'.
        CONTINUE.
      ENDIF.
    ENDIF.

    " Material must be an item of the sales order(s); unit of the item
    lv_found = abap_false.
    IF <ls_upl>-matnr IS NOT INITIAL.
      LOOP AT lt_row_so INTO ls_row_so WHERE idx = lv_idx.
        READ TABLE lt_unit INTO DATA(ls_unit) WITH TABLE KEY vbeln = ls_row_so-vbeln matnr = <ls_upl>-matnr.
        IF sy-subrc = 0.
          lv_found = abap_true.
          <ls_upl>-data-zsd_meins = ls_unit-vrkme.
          EXIT.
        ENDIF.
      ENDLOOP.
      IF lv_found = abap_false.
        lv_text = |Part No. { <ls_upl>-data-zsd_partno } is not an item of sales order { <ls_upl>-data-vbak_vbeln }|.
        PERFORM row_error USING lv_idx 'ZSD_PARTNO' lv_text.
      ENDIF.
    ENDIF.
    IF <ls_upl>-data-zsd_meins IS INITIAL.
      <ls_upl>-data-zsd_meins = gc_meins.
    ENDIF.

    " Quantity without Part No. cannot be checked
    IF <ls_upl>-data-zsd_qty > 0 AND <ls_upl>-matnr IS INITIAL.
      PERFORM row_error USING lv_idx 'ZSD_PARTNO' 'Part No. is required to check the quantity against the delivered quantity'.
      CONTINUE.
    ENDIF.
    IF <ls_upl>-data-zsd_qty = 0 OR lv_found = abap_false.
      CONTINUE.
    ENDIF.

    " Uploaded quantity per sales order(s) + material
    <lv_key> = |{ <ls_upl>-data-vbak_vbeln }/{ <ls_upl>-matnr }|.
    READ TABLE lt_group ASSIGNING FIELD-SYMBOL(<ls_group>) WITH TABLE KEY key = <lv_key>.
    IF sy-subrc <> 0.
      CLEAR ls_group.
      ls_group-key = <lv_key>.
      LOOP AT lt_row_so INTO ls_row_so WHERE idx = lv_idx.
        READ TABLE lt_deliv INTO DATA(ls_deliv) WITH TABLE KEY vbeln = ls_row_so-vbeln matnr = <ls_upl>-matnr.
        IF sy-subrc = 0.
          ls_group-delivered = ls_group-delivered + ls_deliv-qty.
        ENDIF.
      ENDLOOP.
      INSERT ls_group INTO TABLE lt_group ASSIGNING <ls_group>.
    ENDIF.
    <ls_group>-uploaded = <ls_group>-uploaded + <ls_upl>-data-zsd_qty.
  ENDLOOP.

  " ---- Uploaded > delivered -> error on every row of the group --------------
  LOOP AT gt_upl ASSIGNING <ls_upl>.
    lv_idx = sy-tabix.
    READ TABLE lt_rowkey INTO DATA(lv_key) INDEX lv_idx.
    IF lv_key IS INITIAL.
      CONTINUE.
    ENDIF.
    READ TABLE lt_group INTO ls_group WITH TABLE KEY key = lv_key.
    IF sy-subrc = 0 AND ls_group-uploaded > ls_group-delivered.
      lv_text = |{ gc_msg_qty } (material { <ls_upl>-data-zsd_partno }: uploaded { ls_group-uploaded NUMBER = USER }, | &&
                |delivered { ls_group-delivered NUMBER = USER })|.
      PERFORM row_error USING lv_idx 'ZSD_QTY' lv_text.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& FS: gross weight cannot be less than net weight
*&---------------------------------------------------------------------*
FORM check_weights.

  LOOP AT gt_upl INTO DATA(ls_upl).
    DATA(lv_idx) = sy-tabix.
    IF ls_upl-data-zsd_gwt < ls_upl-data-zsd_nwt.
      PERFORM row_error USING lv_idx 'ZSD_GWT' gc_msg_wt.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Per packing list: Smart Form format vs. notifier parties (FS),
*& already saved in ZSD_EPACK_DATA, header values that differ
*&---------------------------------------------------------------------*
FORM check_packing_lists.

  DATA lt_packno TYPE SORTED TABLE OF zsd_epack_data-zsd_packno WITH UNIQUE KEY table_line.
  DATA lt_hdr    TYPE ty_string_tab.
  DATA ls_first  TYPE zsd_epack_data.
  DATA lv_text   TYPE string.
  DATA lv_np     TYPE abap_bool.

  FIELD-SYMBOLS <lv_value> TYPE any.
  FIELD-SYMBOLS <lv_first> TYPE any.

  " Packing lists of the file
  LOOP AT gt_upl INTO DATA(ls_upl) WHERE data-zsd_packno IS NOT INITIAL.
    INSERT ls_upl-data-zsd_packno INTO TABLE lt_packno.
    IF sy-subrc = 0.
      INSERT VALUE #( packno = ls_upl-data-zsd_packno first_row = ls_upl-row ) INTO TABLE gt_pl.
    ENDIF.
  ENDLOOP.
  IF lt_packno IS INITIAL.
    RETURN.
  ENDIF.

  " Already saved (A08): keys only, used for the overwrite
  SELECT zsd_packno, vbak_vbeln, zsd_sn
    FROM zsd_epack_data
    FOR ALL ENTRIES IN @lt_packno
    WHERE zsd_packno = @lt_packno-table_line
    INTO CORRESPONDING FIELDS OF TABLE @gt_old.

  PERFORM header_fields CHANGING lt_hdr.

  LOOP AT gt_pl ASSIGNING FIELD-SYMBOL(<ls_pl>).
    IF line_exists( gt_old[ zsd_packno = <ls_pl>-packno ] ).
      <ls_pl>-exists = abap_true.
      IF p_ovwr = abap_true.
        PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'W'
                              'Packing list is already saved and will be replaced (Overwrite)'.
      ELSE.
        <ls_pl>-error = abap_true.
        PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'E'
                              'Packing list is already saved - tick "Overwrite" to replace it'.
      ENDIF.
    ENDIF.

    " Header values: first row (lowest Sl.No.) with a value; other rows
    " with a different value -> warning (A07)
    CLEAR: ls_first, lv_np.
    LOOP AT gt_upl INTO ls_upl WHERE data-zsd_packno = <ls_pl>-packno.
      DATA(lv_idx) = sy-tabix.
      IF ls_upl-data-zsd_np1 IS NOT INITIAL OR ls_upl-data-zsd_np2 IS NOT INITIAL
         OR ls_upl-data-zsd_np3 IS NOT INITIAL OR ls_upl-data-zsd_np4 IS NOT INITIAL
         OR ls_upl-data-zsd_np5 IS NOT INITIAL.
        IF lv_np = abap_false AND p_wonp = abap_true.
          PERFORM row_error USING lv_idx 'ZSD_NP1'
            'Notifier party maintained - select format "with notifier party" (FS)'.
        ENDIF.
        lv_np = abap_true.
      ENDIF.
      LOOP AT lt_hdr INTO DATA(lv_field).
        ASSIGN COMPONENT lv_field OF STRUCTURE ls_upl-data TO <lv_value>.
        ASSIGN COMPONENT lv_field OF STRUCTURE ls_first    TO <lv_first>.
        IF <lv_value> IS INITIAL.
          CONTINUE.
        ENDIF.
        IF <lv_first> IS INITIAL.
          <lv_first> = <lv_value>.
        ELSEIF <lv_first> <> <lv_value>.
          READ TABLE gt_cols INTO DATA(ls_col) WITH KEY field = lv_field.
          lv_text = |Value differs from the first row of the packing list - the form prints the first value|.
          PERFORM add_msg USING ls_upl-row <ls_pl>-packno ls_col-pos 'W' lv_text.
        ENDIF.
      ENDLOOP.
    ENDLOOP.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Header fields = all columns except item / key columns
*&---------------------------------------------------------------------*
FORM header_fields CHANGING ct_fields TYPE ty_string_tab.

  DATA lt_item TYPE ty_string_tab.

  lt_item = VALUE #( ( `ZSD_PACKNO` ) ( `VBAK_VBELN` ) ( `ZSD_SN` ) ( `ZSD_INV_SN` )
                     ( `ZSD_ARTNO` ) ( `ZSD_STD` ) ( `ZSD_ART` ) ( `ZSD_PARTNO` )
                     ( `ZSD_MATDESC` ) ( `ZSD_QTY` ) ( `ZSD_GWT` ) ( `ZSD_NWT` ) ( `ZSD_HSN` ) ).
  CLEAR ct_fields.
  LOOP AT gt_cols INTO DATA(ls_col) WHERE source <> gc_src_copy.
    IF NOT line_exists( lt_item[ table_line = CONV string( ls_col-field ) ] ).
      APPEND CONV string( ls_col-field ) TO ct_fields.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Packing lists with errors; all-or-nothing: any error -> nothing saved
*&---------------------------------------------------------------------*
FORM decide.

  LOOP AT gt_upl INTO DATA(ls_upl) WHERE error = abap_true AND data-zsd_packno IS NOT INITIAL.
    READ TABLE gt_pl ASSIGNING FIELD-SYMBOL(<ls_pl>) WITH TABLE KEY packno = ls_upl-data-zsd_packno.
    IF sy-subrc = 0.
      <ls_pl>-error = abap_true.
    ENDIF.
  ENDLOOP.

ENDFORM.

FORM any_error CHANGING cv_error TYPE abap_bool.
  cv_error = xsdbool( line_exists( gt_msg[ type = 'E' ] ) ).
ENDFORM.

*&---------------------------------------------------------------------*
*& Save to ZSD_EPACK_DATA
*&   P_ALL : one LUW for all packing lists - saved only without any error
*&   P_PART: one LUW per packing list without error (A09)
*&---------------------------------------------------------------------*
FORM save_data.

  DATA lt_new   TYPE ty_db_tab.
  DATA lt_old   TYPE ty_db_tab.
  DATA lv_error TYPE abap_bool.
  DATA lv_ok    TYPE abap_bool.
  DATA ls_db    TYPE zsd_epack_data.

  PERFORM any_error CHANGING lv_error.
  IF p_all = abap_true AND lv_error = abap_true.
    PERFORM add_msg USING 0 space 0 'E' 'Nothing saved and printed: correct the errors and upload the file again'.
    RETURN.
  ENDIF.

  " Lock per packing list (lock object EZSD_EPACK)
  LOOP AT gt_pl ASSIGNING FIELD-SYMBOL(<ls_pl>) WHERE error = abap_false.
    CALL FUNCTION 'ENQUEUE_EZSD_EPACK'
      EXPORTING
        mode_zsd_epack_data = 'E'
        mandt               = sy-mandt
        zsd_packno          = <ls_pl>-packno
        _scope              = '1'
      EXCEPTIONS
        foreign_lock        = 1
        system_failure      = 2
        OTHERS              = 3.
    IF sy-subrc = 0.
      <ls_pl>-locked = abap_true.
    ELSE.
      <ls_pl>-error = abap_true.
      PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'E'
                            'Packing list is locked by another user - not saved'.
    ENDIF.
  ENDLOOP.

  PERFORM any_error CHANGING lv_error.
  IF p_all = abap_true AND lv_error = abap_true.
    PERFORM add_msg USING 0 space 0 'E' 'Nothing saved and printed: correct the errors and upload the file again'.
    PERFORM unlock.
    RETURN.
  ENDIF.

  LOOP AT gt_pl ASSIGNING <ls_pl> WHERE locked = abap_true.
    " Rows of the packing list, admin fields
    LOOP AT gt_upl INTO DATA(ls_upl) WHERE data-zsd_packno = <ls_pl>-packno.
      ls_db       = ls_upl-data.
      ls_db-mandt = sy-mandt.
      ls_db-ernam = sy-uname.
      ls_db-erdat = sy-datum.
      ls_db-erzet = sy-uzeit.
      APPEND ls_db TO lt_new.
    ENDLOOP.
    IF <ls_pl>-exists = abap_true.
      LOOP AT gt_old INTO ls_db WHERE zsd_packno = <ls_pl>-packno.
        APPEND ls_db TO lt_old.
      ENDLOOP.
    ENDIF.

    IF p_part = abap_true.
      PERFORM write_db USING lt_old lt_new CHANGING lv_ok.
      <ls_pl>-saved = lv_ok.
      PERFORM save_msg USING <ls_pl> lv_ok.
      CLEAR: lt_new, lt_old.
    ENDIF.
  ENDLOOP.

  IF p_all = abap_true AND lt_new IS NOT INITIAL.
    PERFORM write_db USING lt_old lt_new CHANGING lv_ok.
    LOOP AT gt_pl ASSIGNING <ls_pl> WHERE locked = abap_true.
      <ls_pl>-saved = lv_ok.
      PERFORM save_msg USING <ls_pl> lv_ok.
    ENDLOOP.
  ENDIF.

  PERFORM unlock.

ENDFORM.

*&---------------------------------------------------------------------*
*& One LUW: delete the replaced rows, insert the new rows, commit.
*& Any database error -> ROLLBACK WORK, nothing of this LUW is saved.
*&---------------------------------------------------------------------*
FORM write_db USING    it_old TYPE ty_db_tab
                       it_new TYPE ty_db_tab
              CHANGING cv_ok  TYPE abap_bool.

  cv_ok = abap_false.

  IF it_old IS NOT INITIAL.
    DELETE zsd_epack_data FROM TABLE @it_old.
    IF sy-subrc <> 0.
      ROLLBACK WORK.
      RETURN.
    ENDIF.
  ENDIF.

  INSERT zsd_epack_data FROM TABLE @it_new ACCEPTING DUPLICATE KEYS.
  IF sy-subrc <> 0 OR sy-dbcnt <> lines( it_new ).
    ROLLBACK WORK.
    RETURN.
  ENDIF.

  COMMIT WORK AND WAIT.
  cv_ok = abap_true.

ENDFORM.

FORM save_msg USING is_pl TYPE ty_pl
                    iv_ok TYPE abap_bool.
  IF iv_ok = abap_true.
    PERFORM add_msg USING is_pl-first_row is_pl-packno 0 'S' 'Saved in ZSD_EPACK_DATA'.
  ELSE.
    PERFORM add_msg USING is_pl-first_row is_pl-packno 0 'E' 'Database error - not saved (rolled back)'.
  ENDIF.
ENDFORM.

FORM unlock.
  LOOP AT gt_pl ASSIGNING FIELD-SYMBOL(<ls_pl>) WHERE locked = abap_true.
    CALL FUNCTION 'DEQUEUE_EZSD_EPACK'
      EXPORTING
        mode_zsd_epack_data = 'E'
        mandt               = sy-mandt
        zsd_packno          = <ls_pl>-packno
        _scope              = '1'.
    <ls_pl>-locked = abap_false.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Smart Form per saved packing list, all in one spool request
*& (SSF_OPEN / SSF_CLOSE). Format chosen on the selection screen (FS).
*&---------------------------------------------------------------------*
FORM print_forms.

  DATA lv_form  TYPE tdsfname.
  DATA lv_fm    TYPE rs38l_fnam.
  DATA ls_ctrl  TYPE ssfctrlop.
  DATA ls_hdr   TYPE zsd_s_epack_sf_hdr.
  DATA lt_itm   TYPE zsd_tt_epack_sf_itm.
  DATA lv_text  TYPE string.
  DATA lv_wt_ok TYPE abap_bool.

  IF NOT line_exists( gt_pl[ saved = abap_true ] ).
    RETURN.
  ENDIF.

  lv_form = COND #( WHEN p_wnp = abap_true THEN gc_form_np ELSE gc_form_wonp ).
  CALL FUNCTION 'SSF_FUNCTION_MODULE_NAME'
    EXPORTING
      formname           = lv_form
    IMPORTING
      fm_name            = lv_fm
    EXCEPTIONS
      no_form            = 1
      no_function_module = 2
      OTHERS             = 3.
  IF sy-subrc <> 0.
    lv_text = |Smart Form { lv_form } not found - data saved, nothing printed|.
    PERFORM add_msg USING 0 space 0 'E' lv_text.
    RETURN.
  ENDIF.

  CALL FUNCTION 'SSF_OPEN'
    EXPORTING
      user_settings    = abap_true
    EXCEPTIONS
      formatting_error = 1
      internal_error   = 2
      send_error       = 3
      user_canceled    = 4
      OTHERS           = 5.
  IF sy-subrc <> 0.
    PERFORM add_msg USING 0 space 0 'W' 'Printing cancelled - data saved; print later by uploading again with Overwrite'.
    RETURN.
  ENDIF.

  ls_ctrl-no_open  = abap_true.
  ls_ctrl-no_close = abap_true.

  LOOP AT gt_pl ASSIGNING FIELD-SYMBOL(<ls_pl>) WHERE saved = abap_true.
    PERFORM form_data USING <ls_pl>-packno CHANGING ls_hdr lt_itm.

    " FS: no output when a gross weight is less than the net weight
    lv_wt_ok = abap_true.
    LOOP AT lt_itm INTO DATA(ls_itm).
      IF ls_itm-gwt < ls_itm-nwt.
        lv_wt_ok = abap_false.
        EXIT.
      ENDIF.
    ENDLOOP.
    IF lv_wt_ok = abap_false.
      PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'E' gc_msg_wt.
      CONTINUE.
    ENDIF.

    CALL FUNCTION lv_fm
      EXPORTING
        control_parameters = ls_ctrl
        is_header          = ls_hdr
        it_items           = lt_itm
      EXCEPTIONS
        formatting_error   = 1
        internal_error     = 2
        send_error         = 3
        user_canceled      = 4
        OTHERS             = 5.
    IF sy-subrc = 0.
      <ls_pl>-printed = abap_true.
      PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'S' 'Packing list printed'.
    ELSE.
      lv_text = |Smart Form output failed (return code { sy-subrc })|.
      PERFORM add_msg USING <ls_pl>-first_row <ls_pl>-packno 0 'E' lv_text.
    ENDIF.
  ENDLOOP.

  CALL FUNCTION 'SSF_CLOSE'
    EXCEPTIONS
      formatting_error = 1
      internal_error   = 2
      send_error       = 3
      OTHERS           = 4.
  IF sy-subrc <> 0.
    PERFORM add_msg USING 0 space 0 'E' 'Spool request could not be closed'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Smart Form data of one packing list (= saved rows):
*& header from the first row with a value (A07), items by Sl.No.,
*& address columns split into their fixed positions (A14), totals
*&---------------------------------------------------------------------*
FORM form_data USING    iv_packno TYPE zsd_epack_data-zsd_packno
               CHANGING cs_hdr    TYPE zsd_s_epack_sf_hdr
                        ct_itm    TYPE zsd_tt_epack_sf_itm.

  DATA lt_rows  TYPE ty_db_tab.
  DATA lt_hdr   TYPE ty_string_tab.
  DATA ls_h     TYPE zsd_epack_data.
  DATA lt_parts TYPE ty_string_tab.
  DATA lt_so    TYPE SORTED TABLE OF zsd_epack_data-vbak_vbeln WITH UNIQUE KEY table_line.
  DATA lt_split TYPE ty_string_tab.
  DATA lv_text  TYPE string.

  FIELD-SYMBOLS <lv_value> TYPE any.
  FIELD-SYMBOLS <lv_head>  TYPE any.

  CLEAR: cs_hdr, ct_itm.

  LOOP AT gt_upl INTO DATA(ls_upl) WHERE data-zsd_packno = iv_packno.
    APPEND ls_upl-data TO lt_rows.
  ENDLOOP.
  SORT lt_rows BY zsd_sn.

  " Header values: first row with a value
  PERFORM header_fields CHANGING lt_hdr.
  LOOP AT lt_rows INTO DATA(ls_row).
    LOOP AT lt_hdr INTO DATA(lv_field).
      ASSIGN COMPONENT lv_field OF STRUCTURE ls_row TO <lv_value>.
      ASSIGN COMPONENT lv_field OF STRUCTURE ls_h   TO <lv_head>.
      IF <lv_head> IS INITIAL AND <lv_value> IS NOT INITIAL.
        <lv_head> = <lv_value>.
      ENDIF.
    ENDLOOP.
    " Sales orders of the packing list
    lv_text = ls_row-vbak_vbeln.
    SPLIT lv_text AT `,` INTO TABLE lt_split.
    LOOP AT lt_split INTO DATA(lv_so).
      INSERT CONV #( |{ lv_so ALPHA = OUT }| ) INTO TABLE lt_so.
    ENDLOOP.
  ENDLOOP.

  cs_hdr-packno    = iv_packno.
  cs_hdr-packdt    = ls_h-zsd_packdt.
  CLEAR lv_text.
  LOOP AT lt_so INTO DATA(lv_so_out).
    lv_text = COND #( WHEN lv_text IS INITIAL THEN |{ lv_so_out }| ELSE |{ lv_text }, { lv_so_out }| ).
  ENDLOOP.
  cs_hdr-sales_orders = lv_text.

  " Exporter: 1 name, 2 street, 3 street 2, 4 city, 5 postal code, 6 region, 7 country
  PERFORM split_address USING ls_h-zsd_exp 7 CHANGING lt_parts.
  cs_hdr-exp_name     = lt_parts[ 1 ].
  cs_hdr-exp_street   = lt_parts[ 2 ].
  cs_hdr-exp_street2  = lt_parts[ 3 ].
  cs_hdr-exp_city     = lt_parts[ 4 ].
  cs_hdr-exp_postcode = lt_parts[ 5 ].
  cs_hdr-exp_region   = lt_parts[ 6 ].
  cs_hdr-exp_country  = lt_parts[ 7 ].

  cs_hdr-expref    = ls_h-zsd_expref.
  cs_hdr-cin       = ls_h-zsd_cin.
  cs_hdr-gst       = ls_h-zsd_gst.
  cs_hdr-ordno     = ls_h-zsd_ordno.
  cs_hdr-ord_date  = ls_h-zsd_ord_date.
  cs_hdr-party_ref = ls_h-zsd_party_ref.
  cs_hdr-ref_dt    = ls_h-zsd_ref_dt.
  cs_hdr-cnty_orgn = ls_h-zsd_cnty_orgn.
  cs_hdr-cnty_find = ls_h-zsd_cnty_find.

  " Sold-to / ship-to: 1 name, 2 street, 3 street 2, 4 city, 5 district,
  " 6 postal code, 7 country, 8 e-mail, 9 telephone, 10 mobile
  PERFORM split_address USING ls_h-zsd_sold_part 10 CHANGING lt_parts.
  cs_hdr-sold_name     = lt_parts[ 1 ].
  cs_hdr-sold_street   = lt_parts[ 2 ].
  cs_hdr-sold_street2  = lt_parts[ 3 ].
  cs_hdr-sold_city     = lt_parts[ 4 ].
  cs_hdr-sold_district = lt_parts[ 5 ].
  cs_hdr-sold_postcode = lt_parts[ 6 ].
  cs_hdr-sold_country  = lt_parts[ 7 ].
  cs_hdr-sold_email    = lt_parts[ 8 ].
  cs_hdr-sold_tel      = lt_parts[ 9 ].
  cs_hdr-sold_mobile   = lt_parts[ 10 ].

  PERFORM split_address USING ls_h-zsd_ship_part 10 CHANGING lt_parts.
  cs_hdr-ship_name     = lt_parts[ 1 ].
  cs_hdr-ship_street   = lt_parts[ 2 ].
  cs_hdr-ship_street2  = lt_parts[ 3 ].
  cs_hdr-ship_city     = lt_parts[ 4 ].
  cs_hdr-ship_district = lt_parts[ 5 ].
  cs_hdr-ship_postcode = lt_parts[ 6 ].
  cs_hdr-ship_country  = lt_parts[ 7 ].
  cs_hdr-ship_email    = lt_parts[ 8 ].
  cs_hdr-ship_tel      = lt_parts[ 9 ].
  cs_hdr-ship_mobile   = lt_parts[ 10 ].

  " Notifier parties; unused ones are crossed out on the form (FS)
  cs_hdr-np1 = ls_h-zsd_np1.
  cs_hdr-np2 = ls_h-zsd_np2.
  cs_hdr-np3 = ls_h-zsd_np3.
  cs_hdr-np4 = ls_h-zsd_np4.
  cs_hdr-np5 = ls_h-zsd_np5.
  cs_hdr-np_count = COND i( WHEN ls_h-zsd_np1 IS INITIAL THEN 0 ELSE 1 )
                  + COND i( WHEN ls_h-zsd_np2 IS INITIAL THEN 0 ELSE 1 )
                  + COND i( WHEN ls_h-zsd_np3 IS INITIAL THEN 0 ELSE 1 )
                  + COND i( WHEN ls_h-zsd_np4 IS INITIAL THEN 0 ELSE 1 )
                  + COND i( WHEN ls_h-zsd_np5 IS INITIAL THEN 0 ELSE 1 ).

  cs_hdr-prec     = ls_h-zsd_prec.
  cs_hdr-rec_pc   = ls_h-zsd_rec_pc.
  cs_hdr-vsl_flt  = ls_h-zsd_vsl_flt.
  cs_hdr-pol      = ls_h-zsd_pol.
  cs_hdr-pod      = ls_h-zsd_pod.
  cs_hdr-pld      = ls_h-zsd_pld.
  cs_hdr-pterms   = ls_h-zsd_pterms.
  cs_hdr-incot    = ls_h-zsd_incot.
  cs_hdr-mark     = ls_h-zsd_mark.
  cs_hdr-cont     = ls_h-zsd_cont.
  cs_hdr-nopack   = ls_h-zsd_nopack.
  cs_hdr-hsn_desc = ls_h-zsd_hsn_desc.
  cs_hdr-pl_decl  = ls_h-zsd_pl_decl.
  cs_hdr-tart     = ls_h-zsd_tart.
  cs_hdr-gewei    = gc_gewei.

  " Items by Sl.No. and totals
  LOOP AT lt_rows INTO ls_row.
    APPEND VALUE #( sn      = ls_row-zsd_sn
                    inv_sn  = ls_row-zsd_inv_sn
                    artno   = ls_row-zsd_artno
                    std     = ls_row-zsd_std
                    art     = ls_row-zsd_art
                    partno  = ls_row-zsd_partno
                    matdesc = ls_row-zsd_matdesc
                    qty     = ls_row-zsd_qty
                    meins   = ls_row-zsd_meins
                    gwt     = ls_row-zsd_gwt
                    nwt     = ls_row-zsd_nwt
                    gewei   = ls_row-zsd_gewei
                    hsn     = ls_row-zsd_hsn
                    vbeln   = ls_row-vbak_vbeln ) TO ct_itm.
    cs_hdr-tot_art = cs_hdr-tot_art + ls_row-zsd_art.
    cs_hdr-tot_qty = cs_hdr-tot_qty + ls_row-zsd_qty.
    cs_hdr-tot_gwt = cs_hdr-tot_gwt + ls_row-zsd_gwt.
    cs_hdr-tot_nwt = cs_hdr-tot_nwt + ls_row-zsd_nwt.
    cs_hdr-meins   = ls_row-zsd_meins.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Messages
*&---------------------------------------------------------------------*
FORM add_msg USING iv_row    TYPE i
                   iv_packno TYPE csequence
                   iv_pos    TYPE i
                   iv_type   TYPE symsgty
                   iv_text   TYPE csequence.

  DATA lv_letter TYPE string.
  DATA lv_column TYPE string.

  IF iv_pos > 0.
    PERFORM col_letter USING iv_pos CHANGING lv_letter.
    READ TABLE gt_cols INTO DATA(ls_col) WITH KEY pos = iv_pos.
    lv_column = |{ lv_letter } - { ls_col-header }|.
  ENDIF.
  APPEND VALUE #( row = iv_row packno = iv_packno column = lv_column
                  type = iv_type text = iv_text ) TO gt_msg.

ENDFORM.

*&---------------------------------------------------------------------*
*& Error on a row of GT_UPL (index), column given by its field
*&---------------------------------------------------------------------*
FORM row_error USING iv_idx   TYPE i
                     iv_field TYPE csequence
                     iv_text  TYPE csequence.

  DATA lv_pos TYPE i.

  READ TABLE gt_upl ASSIGNING FIELD-SYMBOL(<ls_upl>) INDEX iv_idx.
  IF sy-subrc <> 0.
    RETURN.
  ENDIF.
  <ls_upl>-error = abap_true.
  READ TABLE gt_cols INTO DATA(ls_col) WITH KEY field = iv_field.
  IF sy-subrc = 0.
    lv_pos = ls_col-pos.
  ENDIF.
  PERFORM add_msg USING <ls_upl>-row <ls_upl>-data-zsd_packno lv_pos 'E' iv_text.

ENDFORM.

*&---------------------------------------------------------------------*
*& Result list: summary + all messages with row and column
*&---------------------------------------------------------------------*
FORM show_result.

  DATA lv_errors  TYPE i.
  DATA lv_saved   TYPE i.
  DATA lv_printed TYPE i.
  DATA lv_title   TYPE string.

  LOOP AT gt_msg TRANSPORTING NO FIELDS WHERE type = 'E'.
    lv_errors = lv_errors + 1.
  ENDLOOP.
  LOOP AT gt_pl INTO DATA(ls_pl).
    IF ls_pl-saved = abap_true.
      lv_saved = lv_saved + 1.
    ENDIF.
    IF ls_pl-printed = abap_true.
      lv_printed = lv_printed + 1.
    ENDIF.
  ENDLOOP.

  lv_title = |{ lines( gt_upl ) } rows, { lines( gt_pl ) } packing lists, { lv_errors } errors, | &&
             |{ lv_saved } saved, { lv_printed } printed|.
  IF p_test = abap_true.
    lv_title = |TEST RUN - nothing saved: { lv_title }|.
    IF lv_errors = 0 AND gv_stop = abap_false.
      PERFORM add_msg USING 0 space 0 'S' 'No errors - run again without "Test run" to save and print'.
    ENDIF.
  ENDIF.

  IF gt_msg IS INITIAL.
    MESSAGE lv_title TYPE 'S'.
    RETURN.
  ENDIF.
  PERFORM show_messages USING gt_msg lv_title abap_false.

ENDFORM.

*&---------------------------------------------------------------------*
*& Text elements (Goto -> Text elements)
*&   Text symbols   B01 Upload file
*&                  B02 Smart Form format
*&                  B03 Processing
*&   Selection texts P_FILE  Excel file (.xlsx)
*&                   P_WONP  Packing list without notifier party
*&                   P_WNP   Packing list with notifier party
*&                   P_ALL   Save only if the whole file is correct
*&                   P_PART  Save correct packing lists only
*&                   P_OVWR  Overwrite saved packing list
*&                   P_TEST  Test run (validate only)
*&---------------------------------------------------------------------*
