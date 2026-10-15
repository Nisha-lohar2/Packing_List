*&---------------------------------------------------------------------*
*& Include ZSD_EPACK_COMMON
*&---------------------------------------------------------------------*
*& WRICEF 102-B  Export consolidated packing list
*& Shared by
*&   ZSD_EPACK_EXTRACT  (report 1: sales orders -> editable ALV -> Excel)
*&   ZSD_EPACK_UPLOAD   (report 2: Excel -> validation -> ZSD_EPACK_DATA
*&                       -> Smart Form)
*&
*& The Excel layout is defined ONCE here (FORM epack_columns): columns,
*& headings and order of the sheet "Z table" of Logic_sheet1.xlsx
*& (columns B - AU). Report 1 writes this layout, report 2 expects it.
*&---------------------------------------------------------------------*

CONSTANTS:
  gc_src_sap    TYPE c LENGTH 1 VALUE 'S',   " fetched from SAP - protected in the ALV
  gc_src_manual TYPE c LENGTH 1 VALUE 'M',   " entered by the user - editable
  gc_src_copy   TYPE c LENGTH 1 VALUE 'C',   " repeats another column (sheet column Q)
  gc_max_string TYPE i VALUE 1500.           " ZSD_PL_DECL: CHAR 1500 in the Z-table sheet

TYPES:
  " One column of the Excel file
  BEGIN OF ty_col,
    pos    TYPE i,                  " column in the file, 1 = A
    field  TYPE fieldname,          " field of ZSD_EPACK_DATA
    header TYPE string,             " heading (row 1 of the file)
    source TYPE c LENGTH 1,         " S / M / C, see constants
  END OF ty_col,
  ty_col_tab TYPE STANDARD TABLE OF ty_col WITH EMPTY KEY,

  " Message with its place in the file
  BEGIN OF ty_msg,
    row    TYPE i,                  " Excel row (0 = whole file / packing list)
    packno TYPE c LENGTH 10,        " packing list no.
    column TYPE string,             " Excel column and heading
    type   TYPE symsgty,            " E error, W warning, S success
    text   TYPE string,
  END OF ty_msg,
  ty_msg_tab TYPE STANDARD TABLE OF ty_msg WITH EMPTY KEY,

  ty_string_tab TYPE STANDARD TABLE OF string WITH EMPTY KEY.

*&---------------------------------------------------------------------*
*& Excel layout = sheet "Z table" (Logic_sheet1.xlsx), columns B - AU.
*& Headings as in the sheet (line break replaced by a space).
*& Column 16 (sheet column Q "Packing list no") repeats ZSD_PACKNO: a
*& table cannot hold the same field twice (assumption A04).
*&---------------------------------------------------------------------*
FORM epack_columns CHANGING ct_cols TYPE ty_col_tab.

  ct_cols = VALUE #(
    ( pos =  1 field = 'ZSD_PACKNO'    source = gc_src_manual header = `Packing list no KEY FIELD` )
    ( pos =  2 field = 'VBAK_VBELN'    source = gc_src_sap    header = `SAP Sales Order No KEY FIELD` )
    ( pos =  3 field = 'ZSD_SN'        source = gc_src_manual header = `Sl.No.` )
    ( pos =  4 field = 'ZSD_INV_SN'    source = gc_src_manual header = `Sr.No. INV` )
    ( pos =  5 field = 'ZSD_ARTNO'     source = gc_src_manual header = `Article No.` )
    ( pos =  6 field = 'ZSD_STD'       source = gc_src_manual header = `STD` )
    ( pos =  7 field = 'ZSD_ART'       source = gc_src_manual header = `Articles` )
    ( pos =  8 field = 'ZSD_TART'      source = gc_src_manual header = `Total Articles` )
    ( pos =  9 field = 'ZSD_PARTNO'    source = gc_src_sap    header = `Part No.` )
    ( pos = 10 field = 'ZSD_MATDESC'   source = gc_src_sap    header = `Product Description` )
    ( pos = 11 field = 'ZSD_QTY'       source = gc_src_sap    header = `Quantity      Pcs` )
    ( pos = 12 field = 'ZSD_GWT'       source = gc_src_manual header = `Gross Wt. KGS` )
    ( pos = 13 field = 'ZSD_NWT'       source = gc_src_manual header = `Net Wt. KGS` )
    ( pos = 14 field = 'ZSD_HSN'       source = gc_src_sap    header = `H.S Code` )
    ( pos = 15 field = 'ZSD_EXP'       source = gc_src_sap    header = `Exporter` )
    ( pos = 16 field = 'ZSD_PACKNO'    source = gc_src_copy   header = `Packing list no` )
    ( pos = 17 field = 'ZSD_PACKDT'    source = gc_src_manual header = `Packing list date` )
    ( pos = 18 field = 'ZSD_EXPREF'    source = gc_src_sap    header = `Exporter ReF` )
    ( pos = 19 field = 'ZSD_CIN'       source = gc_src_sap    header = `CIN No.` )
    ( pos = 20 field = 'ZSD_GST'       source = gc_src_sap    header = `GST.No` )
    ( pos = 21 field = 'ZSD_ORDNO'     source = gc_src_sap    header = `Order No` )
    ( pos = 22 field = 'ZSD_ORD_DATE'  source = gc_src_sap    header = `Order date` )
    ( pos = 23 field = 'ZSD_PARTY_REF' source = gc_src_manual header = `Buyer's Reference No` )
    ( pos = 24 field = 'ZSD_REF_DT'    source = gc_src_manual header = `Buyer's Reference date` )
    ( pos = 25 field = 'ZSD_CNTY_ORGN' source = gc_src_manual header = `Country of origin of goods` )
    ( pos = 26 field = 'ZSD_CNTY_FIND' source = gc_src_sap    header = `Country of final destination` )
    ( pos = 27 field = 'ZSD_SOLD_PART' source = gc_src_sap    header = `Sold to party` )
    ( pos = 28 field = 'ZSD_SHIP_PART' source = gc_src_sap    header = `Ship to Party` )
    ( pos = 29 field = 'ZSD_NP1'       source = gc_src_manual header = `Notifier Party1` )
    ( pos = 30 field = 'ZSD_NP2'       source = gc_src_manual header = `Notifier Party2` )
    ( pos = 31 field = 'ZSD_NP3'       source = gc_src_manual header = `Notifier Party3` )
    ( pos = 32 field = 'ZSD_NP4'       source = gc_src_manual header = `Notifier Party4` )
    ( pos = 33 field = 'ZSD_NP5'       source = gc_src_manual header = `Notifier Party5` )
    ( pos = 34 field = 'ZSD_PREC'      source = gc_src_manual header = `Pre-Carriage` )
    ( pos = 35 field = 'ZSD_REC_PC'    source = gc_src_manual header = `Place of Receipt by Pre-carrier` )
    ( pos = 36 field = 'ZSD_POL'       source = gc_src_manual header = `Port of Loading` )
    ( pos = 37 field = 'ZSD_POD'       source = gc_src_manual header = `Port of Discharge` )
    ( pos = 38 field = 'ZSD_PLD'       source = gc_src_manual header = `Place of Delivery` )
    ( pos = 39 field = 'ZSD_PTERMS'    source = gc_src_sap    header = `Terms of Payment` )
    ( pos = 40 field = 'ZSD_INCOT'     source = gc_src_sap    header = `Terms of shipment` )
    ( pos = 41 field = 'ZSD_MARK'      source = gc_src_manual header = `Marks & Nos.` )
    ( pos = 42 field = 'ZSD_VSL_FLT'   source = gc_src_manual header = `Vessel / Flight No.` )
    ( pos = 43 field = 'ZSD_CONT'      source = gc_src_manual header = `Container No` )
    ( pos = 44 field = 'ZSD_NOPACK'    source = gc_src_manual header = `No. of Packages` )
    ( pos = 45 field = 'ZSD_HSN_DESC'  source = gc_src_manual header = `Description of goods` )
    ( pos = 46 field = 'ZSD_PL_DECL'   source = gc_src_manual header = `Declaration of packing list` ) ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Column number -> Excel letter(s): 1 -> A, 27 -> AA
*&---------------------------------------------------------------------*
FORM col_letter USING    iv_pos    TYPE i
                CHANGING cv_letter TYPE string.

  CONSTANTS lc_abc TYPE c LENGTH 26 VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
  DATA lv_num TYPE i.
  DATA lv_rem TYPE i.

  CLEAR cv_letter.
  lv_num = iv_pos.
  WHILE lv_num > 0.
    lv_rem    = ( lv_num - 1 ) MOD 26.
    cv_letter = |{ lc_abc+lv_rem(1) }{ cv_letter }|.
    lv_num    = ( lv_num - 1 ) DIV 26.
  ENDWHILE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Heading -> comparison key (upper case, letters and digits only)
*&---------------------------------------------------------------------*
FORM normalize USING    iv_text TYPE csequence
               CHANGING cv_key  TYPE string.
  cv_key = replace( val = to_upper( iv_text ) regex = `[^A-Z0-9]` with = `` occ = 0 ).
ENDFORM.

*&---------------------------------------------------------------------*
*& Download table: one STRING column per Excel column (C01 ... C46)
*&---------------------------------------------------------------------*
FORM create_xl_table USING    it_cols  TYPE ty_col_tab
                     CHANGING cr_table TYPE REF TO data.

  DATA lt_comp TYPE cl_abap_structdescr=>component_table.
  DATA lo_line TYPE REF TO cl_abap_structdescr.
  DATA lo_tab  TYPE REF TO cl_abap_tabledescr.

  LOOP AT it_cols INTO DATA(ls_col).
    APPEND VALUE #( name = |C{ ls_col-pos WIDTH = 2 ALIGN = RIGHT PAD = '0' }|
                    type = cl_abap_elemdescr=>get_string( ) ) TO lt_comp.
  ENDLOOP.
  lo_line = cl_abap_structdescr=>create( lt_comp ).
  lo_tab  = cl_abap_tabledescr=>create( p_line_type = lo_line ).
  CREATE DATA cr_table TYPE HANDLE lo_tab.

ENDFORM.

*&---------------------------------------------------------------------*
*& Field value -> Excel text. Dates DD.MM.YYYY, numbers with decimal
*& point and no thousands separator, zero / initial -> empty cell.
*&---------------------------------------------------------------------*
FORM field_to_text USING    iv_value TYPE any
                   CHANGING cv_text  TYPE any.

  DATA lv_kind TYPE c LENGTH 1.
  DATA lv_date TYPE d.
  DATA lv_dec  TYPE p LENGTH 16 DECIMALS 3.
  DATA lv_text TYPE string.

  CLEAR cv_text.
  IF iv_value IS INITIAL.
    RETURN.
  ENDIF.

  DESCRIBE FIELD iv_value TYPE lv_kind.
  CASE lv_kind.
    WHEN 'D'.
      lv_date = iv_value.
      lv_text = |{ lv_date+6(2) }.{ lv_date+4(2) }.{ lv_date(4) }|.
    WHEN 'N' OR 'I'.
      lv_text = |{ CONV i( iv_value ) }|.
    WHEN 'P'.
      lv_dec  = iv_value.
      lv_text = |{ lv_dec DECIMALS = 3 }|.
      " 2107.200 -> 2107.2, 60.000 -> 60
      REPLACE REGEX `\.?0+$` IN lv_text WITH ``.
    WHEN OTHERS.
      lv_text = iv_value.
  ENDCASE.
  cv_text = lv_text.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel text -> field of ZSD_EPACK_DATA, checked against the field's
*& type and length. CV_ERROR is filled when the value does not fit.
*&---------------------------------------------------------------------*
FORM text_to_field USING    iv_text  TYPE string
                   CHANGING cv_field TYPE any
                            cv_error TYPE string.

  DATA lv_kind TYPE c LENGTH 1.
  DATA lv_len  TYPE i.
  DATA lv_text TYPE string.
  DATA lv_date TYPE d.
  DATA lv_ok   TYPE abap_bool.
  DATA lv_num  TYPE decfloat34.

  CLEAR: cv_field, cv_error.

  " Leading / trailing blanks and line breaks removed, inner text kept
  lv_text = replace( val = iv_text regex = `^\s+|\s+$` with = `` occ = 0 ).
  IF lv_text IS INITIAL.
    RETURN.
  ENDIF.

  DESCRIBE FIELD cv_field TYPE lv_kind.
  CASE lv_kind.

    WHEN 'D'.                                              " DATS
      PERFORM to_date USING lv_text CHANGING lv_date lv_ok.
      IF lv_ok = abap_true.
        cv_field = lv_date.
      ELSE.
        cv_error = |"{ lv_text }" is not a valid date (use DD.MM.YYYY)|.
      ENDIF.

    WHEN 'N'.                                              " NUMC
      lv_text = condense( val = lv_text del = ` ` ).
      IF NOT matches( val = lv_text regex = `\d+(\.0+)?` ).
        cv_error = |"{ lv_text }" is not a whole number|.
        RETURN.
      ENDIF.
      lv_text = segment( val = lv_text index = 1 sep = `.` ).
      lv_text = shift_left( val = lv_text sub = `0` ).
      DESCRIBE FIELD cv_field LENGTH lv_len IN CHARACTER MODE.
      IF strlen( lv_text ) > lv_len.
        cv_error = |"{ lv_text }" has more than { lv_len } digits|.
        RETURN.
      ENDIF.
      cv_field = lv_text.

    WHEN 'P'.                                              " QUAN
      lv_text = condense( val = lv_text del = ` ` ).
      REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ``.   " thousands separator
      TRY.
          lv_num = lv_text.
          IF lv_num < 0.
            cv_error = |"{ lv_text }" is negative|.
            RETURN.
          ENDIF.
          cv_field = lv_num.                               " rounded to the field's decimals
        CATCH cx_sy_conversion_error cx_sy_arithmetic_error.
          cv_error = |"{ lv_text }" is not a number or too large|.
      ENDTRY.

    WHEN 'g'.                                              " STRING (ZSD_PL_DECL)
      IF strlen( lv_text ) > gc_max_string.
        cv_error = |Text has { strlen( lv_text ) } characters, maximum { gc_max_string }|.
        RETURN.
      ENDIF.
      cv_field = lv_text.

    WHEN OTHERS.                                           " CHAR
      DESCRIBE FIELD cv_field LENGTH lv_len IN CHARACTER MODE.
      IF strlen( lv_text ) > lv_len.
        cv_error = |Text has { strlen( lv_text ) } characters, maximum { lv_len }|.
        RETURN.
      ENDIF.
      cv_field = lv_text.

  ENDCASE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Date text -> date. Accepted: Excel date number, YYYYMMDD,
*& DD.MM.YYYY, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD (time part ignored)
*&---------------------------------------------------------------------*
FORM to_date USING    iv_text TYPE string
             CHANGING cv_date TYPE d
                      cv_ok   TYPE abap_bool.

  DATA lv_text TYPE string.
  DATA lv_day  TYPE string.
  DATA lv_mon  TYPE string.
  DATA lv_year TYPE string.
  DATA lv_num  TYPE decfloat34.
  DATA lv_days TYPE i.

  CLEAR cv_date.
  cv_ok = abap_false.
  lv_text = condense( iv_text ).
  lv_text = segment( val = lv_text index = 1 sep = ` ` ).   " "2026-09-10 00:00:00"

  IF matches( val = lv_text regex = `(19|20)\d{6}` ).
    cv_date = lv_text.                                      " YYYYMMDD
  ELSEIF matches( val = lv_text regex = `\d+(\.\d+)?` ).
    " Excel date number: days since 30.12.1899
    TRY.
        lv_num = lv_text.
      CATCH cx_sy_conversion_error.
        lv_num = 0.
    ENDTRY.
    IF lv_num >= 1 AND lv_num <= 2958465.
      lv_days = trunc( lv_num ).
      cv_date = '18991230'.
      cv_date = cv_date + lv_days.
    ENDIF.
  ELSE.
    REPLACE ALL OCCURRENCES OF REGEX `[/-]` IN lv_text WITH `.`.
    SPLIT lv_text AT '.' INTO lv_day lv_mon lv_year.
    IF strlen( lv_day ) = 4.                                " YYYY-MM-DD
      lv_text = lv_day.
      lv_day  = lv_year.
      lv_year = lv_text.
    ENDIF.
    IF strlen( lv_year ) = 4 AND lv_day CO '0123456789' AND lv_mon CO '0123456789'
       AND lv_year CO '0123456789' AND lv_day IS NOT INITIAL AND lv_mon IS NOT INITIAL
       AND strlen( lv_day ) <= 2 AND strlen( lv_mon ) <= 2.
      cv_date = |{ lv_year }{ lv_mon WIDTH = 2 ALIGN = RIGHT PAD = '0' }{ lv_day WIDTH = 2 ALIGN = RIGHT PAD = '0' }|.
    ENDIF.
  ENDIF.

  IF cv_date IS INITIAL.
    RETURN.
  ENDIF.
  CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
    EXPORTING
      date                      = cv_date
    EXCEPTIONS
      plausibility_check_failed = 1
      OTHERS                    = 2.
  IF sy-subrc = 0.
    cv_ok = abap_true.
  ELSE.
    CLEAR cv_date.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Address parts -> one column with FIXED positions separated by ", ".
*& An empty part stays empty between the commas, so the column can be
*& split back into its parts (FORM split_address). A comma inside a part
*& is replaced by a space. CV_LONG = text did not fit into CV_ADDR.
*&---------------------------------------------------------------------*
FORM fixed_address USING    it_parts TYPE ty_string_tab
                   CHANGING cv_addr  TYPE c
                            cv_long  TYPE abap_bool.

  DATA lt_clean TYPE ty_string_tab.
  DATA lv_text  TYPE string.
  DATA lv_len   TYPE i.

  LOOP AT it_parts INTO lv_text.
    REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ` `.
    CONDENSE lv_text.
    APPEND lv_text TO lt_clean.
  ENDLOOP.
  lv_text = concat_lines_of( table = lt_clean sep = `, ` ).
  cv_addr = lv_text.
  DESCRIBE FIELD cv_addr LENGTH lv_len IN CHARACTER MODE.
  cv_long = xsdbool( strlen( lv_text ) > lv_len ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Fixed-position column -> its parts (always IV_COUNT parts)
*&---------------------------------------------------------------------*
FORM split_address USING    iv_addr  TYPE csequence
                            iv_count TYPE i
                   CHANGING ct_parts TYPE ty_string_tab.

  DATA lv_text TYPE string.

  CLEAR ct_parts.
  lv_text = iv_addr.
  IF lv_text IS NOT INITIAL.
    SPLIT lv_text AT `,` INTO TABLE ct_parts.
  ENDIF.
  LOOP AT ct_parts ASSIGNING FIELD-SYMBOL(<lv_part>).
    CONDENSE <lv_part>.
  ENDLOOP.
  WHILE lines( ct_parts ) < iv_count.
    APPEND INITIAL LINE TO ct_parts.
  ENDWHILE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Messages as ALV: popup (IV_POPUP = X) or full screen
*&---------------------------------------------------------------------*
FORM show_messages USING it_msg   TYPE ty_msg_tab
                         iv_title TYPE csequence
                         iv_popup TYPE abap_bool.

  DATA lt_msg  TYPE ty_msg_tab.
  DATA lo_alv  TYPE REF TO cl_salv_table.
  DATA lx_salv TYPE REF TO cx_salv_error.
  DATA lv_l    TYPE scrtext_l.
  DATA lv_m    TYPE scrtext_m.
  DATA lv_s    TYPE scrtext_s.
  DATA lv_ttl  TYPE lvc_title.
  DATA lt_names TYPE ty_string_tab.

  IF it_msg IS INITIAL.
    RETURN.
  ENDIF.
  lt_msg = it_msg.
  SORT lt_msg BY row packno type column.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv
                              CHANGING  t_table      = lt_msg ).
      IF iv_popup = abap_true.
        lo_alv->set_screen_popup( start_column = 5  end_column = 160
                                  start_line   = 3  end_line   = 22 ).
      ELSE.
        lo_alv->get_functions( )->set_all( abap_true ).
      ENDIF.
      lv_ttl = iv_title.
      lo_alv->get_display_settings( )->set_list_header( lv_ttl ).
      lo_alv->get_display_settings( )->set_striped_pattern( abap_true ).
      lo_alv->get_columns( )->set_optimize( abap_true ).
      lt_names = VALUE #( ( `ROW` ) ( `PACKNO` ) ( `COLUMN` ) ( `TYPE` ) ( `TEXT` ) ).
      LOOP AT lt_names INTO DATA(lv_name).
        CASE lv_name.
          WHEN `ROW`.    lv_l = 'Excel row'.
          WHEN `PACKNO`. lv_l = 'Packing list no'.
          WHEN `COLUMN`. lv_l = 'Column'.
          WHEN `TYPE`.   lv_l = 'Type'.
          WHEN `TEXT`.   lv_l = 'Message'.
        ENDCASE.
        lv_m = lv_l.
        lv_s = lv_l.
        DATA(lo_col) = lo_alv->get_columns( )->get_column( CONV lvc_fname( lv_name ) ).
        lo_col->set_long_text( lv_l ).
        lo_col->set_medium_text( lv_m ).
        lo_col->set_short_text( lv_s ).
      ENDLOOP.
      lo_alv->display( ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& F4 help for an Excel file
*&---------------------------------------------------------------------*
FORM f4_file CHANGING cv_file TYPE rlgrap-filename.

  DATA lt_files TYPE filetable.
  DATA lv_rc    TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      file_filter    = 'Excel (*.xlsx)|*.xlsx'
      multiselection = abap_false
    CHANGING
      file_table     = lt_files
      rc             = lv_rc
    EXCEPTIONS
      OTHERS         = 1 ).
  IF sy-subrc = 0 AND lv_rc = 1.
    cv_file = lt_files[ 1 ]-filename.
  ENDIF.

ENDFORM.
