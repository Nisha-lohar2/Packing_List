"! <p class="shorttext synchronized">Export Packing List - business rules</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! All rules that do NOT need the database: package grouping, the FS
"! weight validation, totals, output formatting and Excel value parsing.
"! Kept free of SELECTs so that every rule is covered by ABAP Unit
"! (see the local test class).
"!
"! PACKAGE GROUPS (not described in the FS - derived from the sample
"! outputs and Z table Format1.xlsx, see ASSUMPTION A06)
"!   In the approved Excel layout "Articles" and "Gross/Net Wt." are
"!   merged cells spanning several material lines. Excel stores a merged
"!   value only in the first cell, so:
"!     - a line with Articles filled starts a new ARTICLE group,
"!       lines with Articles empty belong to the group above;
"!     - a line with Gross/Net weight filled starts a new WEIGHT group,
"!       lines with both weights empty belong to the group above.
"!   The two groupings are independent: in sample 90000014 every line has
"!   its own Articles while three lines share one weight.
"!   Articles / weights are stored only on the first line of a group, so
"!   plain column sums give the totals of the samples exactly
"!   (90000081: 4074 / 12146 / 116323.800 / 115386.780).
CLASS zcl_sd_epack_rules DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "! Build one BAPIRET2 line for message class ZSD_EPACK
    CLASS-METHODS msg
      IMPORTING iv_type          TYPE bapi_mtype
                iv_no            TYPE symsgno
                iv_v1            TYPE simple OPTIONAL
                iv_v2            TYPE simple OPTIONAL
                iv_v3            TYPE simple OPTIONAL
                iv_v4            TYPE simple OPTIONAL
      RETURNING VALUE(rs_return) TYPE bapiret2.

    CLASS-METHODS has_errors
      IMPORTING it_messages      TYPE bapiret2_t
      RETURNING VALUE(rv_errors) TYPE abap_bool.

    "! Excel lines -> table rows with article / weight groups
    "! (errors 021, 027, 028, 041 are appended to CT_MESSAGES)
    CLASS-METHODS build_items
      IMPORTING iv_packno   TYPE zsd_packno
                it_lines    TYPE zif_sd_epack=>tt_upload_line
      EXPORTING et_items    TYPE zif_sd_epack=>tt_item
      CHANGING  ct_messages TYPE bapiret2_t.

    "! FS 2.1: "Gross Weight is always equal to or greater than the Net
    "! Weight ... error message: Invalid Weight: Gross Weight cannot be
    "! less than Net Weight." Checked per weight group (ASSUMPTION A07);
    "! the grand total follows automatically.
    CLASS-METHODS check_weights
      IMPORTING it_items           TYPE zif_sd_epack=>tt_item
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

    "! Quantity = STD x Articles holds for every line of both samples.
    "! Not an FS rule - reported as WARNING only (ASSUMPTION A08).
    CLASS-METHODS check_quantities
      IMPORTING it_items           TYPE zif_sd_epack=>tt_item
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

    CLASS-METHODS calc_totals
      IMPORTING it_items         TYPE zif_sd_epack=>tt_item
      RETURNING VALUE(rs_totals) TYPE zif_sd_epack=>ts_totals.

    "! "Total 4074 Bundles" / "Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages"
    CLASS-METHODS total_text
      IMPORTING is_totals      TYPE zif_sd_epack=>ts_totals
      RETURNING VALUE(rv_text) TYPE string.

    "! Items (HSN and description already resolved) -> printed rows with
    "! HSN section headings and package-type subtotals
    CLASS-METHODS build_print_items
      IMPORTING it_items        TYPE zif_sd_epack=>tt_item
                it_headings     TYPE zif_sd_epack=>tt_hsn_heading
      RETURNING VALUE(rt_print) TYPE zsd_tt_epack_prt_itm.

    "! Section heading for one HSN: the "Description of Goods" line that
    "! contains the HSN, otherwise "<T604N text> - HS CODE : <hsn>"
    CLASS-METHODS heading_for_hsn
      IMPORTING iv_hsn         TYPE zsd_hsn
                it_goods_lines TYPE zif_sd_epack=>tt_text_line
                iv_hsn_text    TYPE string OPTIONAL
      RETURNING VALUE(rv_text) TYPE string.

    "! "39172390" -> "3917 2390" (format of the samples)
    CLASS-METHODS format_hsn
      IMPORTING iv_hsn        TYPE csequence
      RETURNING VALUE(rv_hsn) TYPE zsd_hsn.

    CLASS-METHODS format_weight
      IMPORTING iv_weight      TYPE zsd_gwt
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS format_quantity
      IMPORTING iv_quantity    TYPE zsd_qty
      RETURNING VALUE(rv_text) TYPE string.

    "! DD-MM-YYYY, as on the samples ("Date : 27-05-2025")
    CLASS-METHODS format_date
      IMPORTING iv_date        TYPE d
      RETURNING VALUE(rv_text) TYPE string.

    "! Split a multi-line Excel cell (Alt+Enter) into lines
    CLASS-METHODS split_lines
      IMPORTING iv_text         TYPE csequence
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    "! Excel cell -> number. Accepts "2107.2000000000003", "1,234.5".
    CLASS-METHODS try_parse_number
      IMPORTING iv_text      TYPE csequence
      EXPORTING ev_value     TYPE decfloat34
      RETURNING VALUE(rv_ok) TYPE abap_bool.

    "! Excel cell -> date. Accepts an Excel serial number (45804),
    "! DD.MM.YYYY / DD-MM-YYYY / DD/MM/YYYY, YYYY-MM-DD and YYYYMMDD.
    CLASS-METHODS try_parse_date
      IMPORTING iv_text      TYPE csequence
      EXPORTING ev_date      TYPE d
      RETURNING VALUE(rv_ok) TYPE abap_bool.

    "! "90000014, 90000015 & 90000020" -> invoice numbers (ALPHA format)
    CLASS-METHODS split_invoices
      IMPORTING iv_text    TYPE csequence
      EXPORTING et_vbeln   TYPE zif_sd_epack=>tt_vbeln
                et_invalid TYPE string_table.

  PRIVATE SECTION.

    CLASS-METHODS digits_only
      IMPORTING iv_text        TYPE csequence
      RETURNING VALUE(rv_text) TYPE string.

ENDCLASS.


CLASS zcl_sd_epack_rules IMPLEMENTATION.

  METHOD msg.

    rs_return = VALUE #( type       = iv_type
                         id         = zif_sd_epack=>gc_msgid
                         number     = iv_no
                         message_v1 = |{ iv_v1 }|
                         message_v2 = |{ iv_v2 }|
                         message_v3 = |{ iv_v3 }|
                         message_v4 = |{ iv_v4 }| ).

    MESSAGE ID zif_sd_epack=>gc_msgid TYPE 'S' NUMBER iv_no
            WITH rs_return-message_v1 rs_return-message_v2
                 rs_return-message_v3 rs_return-message_v4
            INTO rs_return-message.

  ENDMETHOD.


  METHOD has_errors.

    LOOP AT it_messages TRANSPORTING NO FIELDS WHERE type CA 'EAX'.
      rv_errors = abap_true.
      RETURN.
    ENDLOOP.

  ENDMETHOD.


  METHOD build_items.

    DATA lv_art_grp  TYPE zsd_art_grp.
    DATA lv_wt_grp   TYPE zsd_wt_grp.
    DATA lv_pkg_type TYPE zsd_pkg_type.
    DATA lv_prev_sn  TYPE zsd_sn.
    DATA lt_seen_sn  TYPE SORTED TABLE OF zsd_sn WITH UNIQUE KEY table_line.

    CLEAR et_items.

    LOOP AT it_lines INTO DATA(ls_line).

      " Sl.No. must be unique and ascending - the groups depend on the
      " row order of the file, the printout is sorted by Sl.No.
      INSERT ls_line-sn INTO TABLE lt_seen_sn.
      IF sy-subrc <> 0.
        APPEND msg( iv_type = 'E' iv_no = '021' iv_v1 = ls_line-row iv_v2 = ls_line-sn )
               TO ct_messages.
      ELSEIF ls_line-sn < lv_prev_sn.
        APPEND msg( iv_type = 'E' iv_no = '041' iv_v1 = ls_line-row iv_v2 = ls_line-sn )
               TO ct_messages.
      ENDIF.
      lv_prev_sn = ls_line-sn.

      " Article group
      IF ls_line-art_filled = abap_true.
        lv_art_grp += 1.
      ELSEIF lv_art_grp IS INITIAL.
        APPEND msg( iv_type = 'E' iv_no = '027' iv_v1 = ls_line-row iv_v2 = 'Articles' )
               TO ct_messages.
      ENDIF.

      " Weight group - gross and net always come together
      IF ls_line-gwt_filled = abap_true OR ls_line-nwt_filled = abap_true.
        IF ls_line-gwt_filled = abap_false OR ls_line-nwt_filled = abap_false.
          APPEND msg( iv_type = 'E' iv_no = '028' iv_v1 = ls_line-row )
                 TO ct_messages.
        ENDIF.
        lv_wt_grp += 1.
      ELSEIF lv_wt_grp IS INITIAL.
        APPEND msg( iv_type = 'E' iv_no = '027' iv_v1 = ls_line-row iv_v2 = 'Gross/Net Wt.' )
               TO ct_messages.
      ENDIF.

      " Package type is entered once per run of lines (ASSUMPTION A09)
      IF ls_line-pkg_type IS NOT INITIAL.
        lv_pkg_type = ls_line-pkg_type.
      ENDIF.

      APPEND VALUE #( zsd_packno   = iv_packno
                      zsd_sn       = ls_line-sn
                      zsd_inv_sn   = ls_line-inv_sn
                      zsd_artno    = ls_line-artno
                      zsd_std      = ls_line-std
                      zsd_art      = COND #( WHEN ls_line-art_filled = abap_true THEN ls_line-art )
                      zsd_art_grp  = lv_art_grp
                      zsd_wt_grp   = lv_wt_grp
                      zsd_pkg_type = lv_pkg_type
                      zsd_partno   = ls_line-partno
                      zsd_matdesc  = ls_line-matdesc
                      zsd_qty      = ls_line-qty
                      zsd_qty_uom  = zif_sd_epack=>gc_uom-quantity
                      zsd_gwt      = COND #( WHEN ls_line-gwt_filled = abap_true THEN ls_line-gwt )
                      zsd_nwt      = COND #( WHEN ls_line-nwt_filled = abap_true THEN ls_line-nwt )
                      zsd_wt_uom   = zif_sd_epack=>gc_uom-weight
                      zsd_hsn      = ls_line-hsn )
             TO et_items.

    ENDLOOP.

  ENDMETHOD.


  METHOD check_weights.

    DATA lv_prev_grp TYPE zsd_wt_grp.
    DATA lt_detail   TYPE bapiret2_t.

    LOOP AT it_items INTO DATA(ls_item).
      " Weights are stored on the first line of each weight group only
      IF ls_item-zsd_wt_grp = lv_prev_grp.
        CONTINUE.
      ENDIF.
      lv_prev_grp = ls_item-zsd_wt_grp.

      IF ls_item-zsd_gwt < ls_item-zsd_nwt.
        APPEND msg( iv_type = 'E'
                    iv_no   = '002'
                    iv_v1   = |{ CONV i( ls_item-zsd_sn ) }|
                    iv_v2   = format_weight( ls_item-zsd_gwt )
                    iv_v3   = format_weight( ls_item-zsd_nwt ) )
               TO lt_detail.
      ENDIF.
    ENDLOOP.

    IF lt_detail IS NOT INITIAL.
      " The exact FS text first, then one detail line per failing group
      APPEND msg( iv_type = 'E' iv_no = '001' ) TO rt_messages.
      APPEND LINES OF lt_detail TO rt_messages.
    ENDIF.

  ENDMETHOD.


  METHOD check_quantities.

    DATA lv_articles TYPE zsd_art.
    DATA lv_prev_grp TYPE zsd_art_grp.

    LOOP AT it_items INTO DATA(ls_item).
      IF ls_item-zsd_art_grp <> lv_prev_grp.
        lv_articles = ls_item-zsd_art.
        lv_prev_grp = ls_item-zsd_art_grp.
      ENDIF.

      DATA(lv_expected) = CONV zsd_qty( ls_item-zsd_std * lv_articles ).
      IF lv_expected <> ls_item-zsd_qty.
        APPEND msg( iv_type = 'W'
                    iv_no   = '029'
                    iv_v1   = |{ CONV i( ls_item-zsd_sn ) }|
                    iv_v2   = format_quantity( ls_item-zsd_qty )
                    iv_v3   = |{ CONV i( ls_item-zsd_std ) }|
                    iv_v4   = |{ CONV i( lv_articles ) }| )
               TO rt_messages.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD calc_totals.

    LOOP AT it_items INTO DATA(ls_item).
      rs_totals-articles += ls_item-zsd_art.
      rs_totals-quantity += ls_item-zsd_qty.
      rs_totals-gross    += ls_item-zsd_gwt.
      rs_totals-net      += ls_item-zsd_nwt.

      ASSIGN rs_totals-pkg_types[ pkg_type = ls_item-zsd_pkg_type ] TO FIELD-SYMBOL(<ls_pkg>).
      IF sy-subrc <> 0.
        APPEND VALUE #( pkg_type = ls_item-zsd_pkg_type ) TO rs_totals-pkg_types
               ASSIGNING <ls_pkg>.
      ENDIF.
      <ls_pkg>-articles += ls_item-zsd_art.
    ENDLOOP.

  ENDMETHOD.


  METHOD total_text.

    DATA(lt_types) = is_totals-pkg_types.
    DELETE lt_types WHERE pkg_type IS INITIAL AND articles = 0.

    CASE lines( lt_types ).
      WHEN 0.
        rv_text = |Total { is_totals-articles } Packages|.
      WHEN 1.
        rv_text = |Total { is_totals-articles } {
                    COND string( WHEN lt_types[ 1 ]-pkg_type IS INITIAL
                                 THEN `Packages`
                                 ELSE condense( CONV string( lt_types[ 1 ]-pkg_type ) ) ) }|.
      WHEN OTHERS.
        rv_text = `Total `.
        LOOP AT lt_types INTO DATA(ls_type).
          rv_text = |{ rv_text }{ COND string( WHEN sy-tabix > 1 THEN ` + ` ) }{ ls_type-articles } {
                      COND string( WHEN ls_type-pkg_type IS INITIAL
                                   THEN `Packages`
                                   ELSE condense( CONV string( ls_type-pkg_type ) ) ) }|.
        ENDLOOP.
        rv_text = |{ rv_text } = { is_totals-articles } Packages|.
    ENDCASE.

  ENDMETHOD.


  METHOD build_print_items.

    DATA lv_prev_hsn   TYPE zsd_hsn VALUE '#'.
    DATA lv_prev_art   TYPE zsd_art_grp.
    DATA lv_prev_wt    TYPE zsd_wt_grp.
    DATA lv_run_type   TYPE zsd_pkg_type.
    DATA lv_run_sum    TYPE i.

    DATA(lt_items) = it_items.
    SORT lt_items BY zsd_sn.

    " Subtotal rows only when more than one package type is used
    " (sample 90000014: Loose Pipes / Boxes; 90000081 has none)
    DATA(lv_subtotals) = xsdbool( lines( calc_totals( lt_items )-pkg_types ) > 1 ).

    LOOP AT lt_items INTO DATA(ls_item).

      IF lv_subtotals = abap_true AND sy-tabix > 1 AND ls_item-zsd_pkg_type <> lv_run_type.
        APPEND VALUE #( row_type = zif_sd_epack=>gc_row_type-subtotal
                        art_txt  = |{ lv_run_sum }|
                        row_text = |<={ condense( CONV string( lv_run_type ) ) }| )
               TO rt_print.
        CLEAR lv_run_sum.
      ENDIF.

      IF ls_item-zsd_hsn <> lv_prev_hsn.
        APPEND VALUE #( row_type = zif_sd_epack=>gc_row_type-heading
                        row_text = VALUE #( it_headings[ hsn = ls_item-zsd_hsn ]-text
                                            DEFAULT |HS CODE : { ls_item-zsd_hsn }| ) )
               TO rt_print.
        lv_prev_hsn = ls_item-zsd_hsn.
      ENDIF.

      DATA(lv_art_first) = xsdbool( ls_item-zsd_art_grp <> lv_prev_art ).
      DATA(lv_wt_first)  = xsdbool( ls_item-zsd_wt_grp  <> lv_prev_wt ).

      APPEND VALUE #( row_type   = zif_sd_epack=>gc_row_type-item
                      sn_txt     = |{ CONV i( ls_item-zsd_sn ) }|
                      inv_sn_txt = COND #( WHEN ls_item-zsd_inv_sn IS NOT INITIAL
                                           THEN |{ CONV i( ls_item-zsd_inv_sn ) }| )
                      artno      = ls_item-zsd_artno
                      std_txt    = |{ CONV i( ls_item-zsd_std ) }|
                      art_txt    = COND #( WHEN lv_art_first = abap_true
                                           THEN |{ CONV i( ls_item-zsd_art ) }| )
                      partno     = ls_item-zsd_partno
                      matdesc    = ls_item-zsd_matdesc
                      qty_txt    = format_quantity( ls_item-zsd_qty )
                      gwt_txt    = COND #( WHEN lv_wt_first = abap_true
                                           THEN format_weight( ls_item-zsd_gwt ) )
                      nwt_txt    = COND #( WHEN lv_wt_first = abap_true
                                           THEN format_weight( ls_item-zsd_nwt ) )
                      hsn        = ls_item-zsd_hsn
                      art_first  = lv_art_first
                      wt_first   = lv_wt_first
                      art_grp    = ls_item-zsd_art_grp
                      wt_grp     = ls_item-zsd_wt_grp )
             TO rt_print.

      lv_prev_art = ls_item-zsd_art_grp.
      lv_prev_wt  = ls_item-zsd_wt_grp.
      lv_run_type = ls_item-zsd_pkg_type.
      lv_run_sum += ls_item-zsd_art.

    ENDLOOP.

    IF lv_subtotals = abap_true AND lt_items IS NOT INITIAL.
      APPEND VALUE #( row_type = zif_sd_epack=>gc_row_type-subtotal
                      art_txt  = |{ lv_run_sum }|
                      row_text = |<={ condense( CONV string( lv_run_type ) ) }| )
             TO rt_print.
    ENDIF.

    " Second pass: last line of each group - the form uses ART_LAST /
    " WT_LAST to draw the bottom border of the "merged" cells.
    DATA lv_next TYPE i.
    DATA ls_next TYPE zsd_s_epack_prt_itm.
    LOOP AT rt_print ASSIGNING FIELD-SYMBOL(<ls_row>)
         WHERE row_type = zif_sd_epack=>gc_row_type-item.
      lv_next = sy-tabix + 1.
      ls_next = VALUE #( rt_print[ lv_next ] OPTIONAL ).
      <ls_row>-art_last = xsdbool( ls_next-row_type <> zif_sd_epack=>gc_row_type-item
                                   OR ls_next-art_grp <> <ls_row>-art_grp ).
      <ls_row>-wt_last  = xsdbool( ls_next-row_type <> zif_sd_epack=>gc_row_type-item
                                   OR ls_next-wt_grp <> <ls_row>-wt_grp ).
    ENDLOOP.

  ENDMETHOD.


  METHOD heading_for_hsn.

    DATA(lv_hsn_digits) = digits_only( iv_hsn ).

    IF lv_hsn_digits IS NOT INITIAL.
      LOOP AT it_goods_lines INTO DATA(lv_line).
        IF digits_only( lv_line ) CS lv_hsn_digits.
          rv_text = lv_line.
          RETURN.
        ENDIF.
      ENDLOOP.
    ENDIF.

    rv_text = COND #( WHEN iv_hsn_text IS INITIAL
                      THEN |HS CODE : { format_hsn( iv_hsn ) }|
                      ELSE |{ iv_hsn_text } - HS CODE : { format_hsn( iv_hsn ) }| ).

  ENDMETHOD.


  METHOD format_hsn.

    DATA(lv_text) = replace( val = condense( CONV string( iv_hsn ) ) sub = ` ` with = `` occ = 0 ).

    IF strlen( lv_text ) = 8 AND lv_text CO '0123456789'.
      rv_hsn = |{ lv_text(4) } { lv_text+4(4) }|.
    ELSE.
      rv_hsn = condense( CONV string( iv_hsn ) ).
    ENDIF.

  ENDMETHOD.


  METHOD format_weight.
    rv_text = |{ iv_weight DECIMALS = 3 }|.
  ENDMETHOD.


  METHOD format_quantity.

    IF frac( iv_quantity ) = 0.
      rv_text = |{ CONV i( iv_quantity ) }|.
    ELSE.
      rv_text = |{ iv_quantity DECIMALS = 3 }|.
    ENDIF.

  ENDMETHOD.


  METHOD format_date.

    IF iv_date IS NOT INITIAL.
      rv_text = |{ iv_date+6(2) }-{ iv_date+4(2) }-{ iv_date(4) }|.
    ENDIF.

  ENDMETHOD.


  METHOD split_lines.

    DATA(lv_text) = CONV string( iv_text ).
    DATA(lv_cr)   = substring( val = cl_abap_char_utilities=>cr_lf len = 1 ).
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN lv_text WITH cl_abap_char_utilities=>newline.
    REPLACE ALL OCCURRENCES OF lv_cr IN lv_text WITH cl_abap_char_utilities=>newline.

    SPLIT lv_text AT cl_abap_char_utilities=>newline INTO TABLE DATA(lt_raw).

    LOOP AT lt_raw INTO DATA(lv_line).
      lv_line = condense( val = lv_line del = ` ` ).
      IF lv_line IS NOT INITIAL.
        APPEND lv_line TO rt_lines.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD try_parse_number.

    CLEAR ev_value.

    DATA(lv_text) = condense( val = CONV string( iv_text ) del = ` ` ).
    REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ``.

    IF lv_text IS INITIAL.
      rv_ok = abap_true.
      RETURN.
    ENDIF.

    TRY.
        ev_value = CONV decfloat34( lv_text ).
        rv_ok = abap_true.
      CATCH cx_sy_conversion_error.
        rv_ok = abap_false.
    ENDTRY.

  ENDMETHOD.


  METHOD try_parse_date.

    DATA lv_date TYPE d.
    DATA lv_day  TYPE string.
    DATA lv_mon  TYPE string.
    DATA lv_year TYPE string.

    CLEAR ev_date.
    DATA(lv_text) = condense( val = CONV string( iv_text ) del = ` ` ).

    IF lv_text IS INITIAL.
      rv_ok = abap_false.
      RETURN.
    ENDIF.

    IF strlen( lv_text ) = 8 AND lv_text CO '0123456789' AND ( lv_text(2) = '19' OR lv_text(2) = '20' ).
      " YYYYMMDD
      lv_date = lv_text.
    ELSEIF lv_text CO '0123456789.'.
      " Excel serial date: day 1 = 01.01.1900 (Excel counts 29.02.1900,
      " hence the base date 30.12.1899)
      DATA lv_serial TYPE decfloat34.
      IF try_parse_number( EXPORTING iv_text = lv_text IMPORTING ev_value = lv_serial ) = abap_false
         OR lv_serial < 1 OR lv_serial > 2958465.
        RETURN.
      ENDIF.
      lv_date = '18991230'.
      lv_date = lv_date + CONV i( trunc( lv_serial ) ).
    ELSEIF matches( val = lv_text regex = `\d{4}-\d{1,2}-\d{1,2}` ).
      SPLIT lv_text AT '-' INTO lv_year lv_mon lv_day.
      lv_date = |{ lv_year }{ lv_mon WIDTH = 2 ALIGN = RIGHT PAD = '0' }{ lv_day WIDTH = 2 ALIGN = RIGHT PAD = '0' }|.
    ELSEIF matches( val = lv_text regex = `\d{1,2}[./-]\d{1,2}[./-]\d{4}` ).
      REPLACE ALL OCCURRENCES OF REGEX `[/-]` IN lv_text WITH `.`.
      SPLIT lv_text AT '.' INTO lv_day lv_mon lv_year.
      lv_date = |{ lv_year }{ lv_mon WIDTH = 2 ALIGN = RIGHT PAD = '0' }{ lv_day WIDTH = 2 ALIGN = RIGHT PAD = '0' }|.
    ELSE.
      RETURN.
    ENDIF.

    CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
      EXPORTING
        date                      = lv_date
      EXCEPTIONS
        plausibility_check_failed = 1
        OTHERS                    = 2.
    IF sy-subrc = 0.
      ev_date = lv_date.
      rv_ok   = abap_true.
    ENDIF.

  ENDMETHOD.


  METHOD split_invoices.

    CLEAR: et_vbeln, et_invalid.

    DATA(lv_text) = CONV string( iv_text ).
    REPLACE ALL OCCURRENCES OF REGEX `[,;&/\s]+` IN lv_text WITH ` `.
    SPLIT condense( lv_text ) AT ` ` INTO TABLE DATA(lt_token).

    LOOP AT lt_token INTO DATA(lv_token) WHERE table_line IS NOT INITIAL.
      IF lv_token CO '0123456789' AND strlen( lv_token ) <= 10.
        INSERT CONV vbeln_vf( |{ lv_token ALPHA = IN WIDTH = 10 }| ) INTO TABLE et_vbeln.
      ELSE.
        APPEND lv_token TO et_invalid.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD digits_only.
    rv_text = replace( val = iv_text regex = `[^0-9]` with = `` occ = 0 ).
  ENDMETHOD.

ENDCLASS.
