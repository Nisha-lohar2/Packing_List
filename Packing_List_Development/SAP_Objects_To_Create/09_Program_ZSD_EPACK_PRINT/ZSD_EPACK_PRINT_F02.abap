*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_F02
*&---------------------------------------------------------------------*
*& Helper routines without database access: messages, package groups,
*& FS weight rule, totals, printed rows, formatting and parsing.
*&
*& PACKAGE GROUPS (A06) - not described in the FS, derived from the
*& sample outputs, Z table Format1.xlsx and Sample_data_ABAP.xlsx:
*&   - a line with Articles filled starts a new ARTICLE group, a line with
*&     Gross/Net Wt. filled starts a new WEIGHT group; empty cells (merged
*&     in Excel) belong to the group above;
*&   - values REPEATED on each row with the same Article No. are the same
*&     package (FORM collapse_repeated_groups).
*&   Articles / weights are stored on the first line of a group only, so
*&   plain column sums give the totals of the samples exactly.
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& Append one message of class ZSD_EPACK
*&---------------------------------------------------------------------*
FORM add_msg USING    iv_type     TYPE simple
                      iv_no       TYPE simple
                      iv_v1       TYPE simple
                      iv_v2       TYPE simple
                      iv_v3       TYPE simple
                      iv_v4       TYPE simple
             CHANGING ct_messages TYPE bapiret2_t.

  DATA ls_msg TYPE bapiret2.

  ls_msg-type       = iv_type.
  ls_msg-id         = gc_msgid.
  ls_msg-number     = iv_no.
  ls_msg-message_v1 = |{ iv_v1 }|.
  ls_msg-message_v2 = |{ iv_v2 }|.
  ls_msg-message_v3 = |{ iv_v3 }|.
  ls_msg-message_v4 = |{ iv_v4 }|.

  MESSAGE ID gc_msgid TYPE 'S' NUMBER ls_msg-number
          WITH ls_msg-message_v1 ls_msg-message_v2
               ls_msg-message_v3 ls_msg-message_v4
          INTO ls_msg-message.

  APPEND ls_msg TO ct_messages.

ENDFORM.

*&---------------------------------------------------------------------*
*& X if the message table contains an error
*&---------------------------------------------------------------------*
FORM has_errors USING    it_messages TYPE bapiret2_t
                CHANGING cv_errors   TYPE abap_bool.

  cv_errors = abap_false.
  LOOP AT it_messages TRANSPORTING NO FIELDS WHERE type CA 'EAX'.
    cv_errors = abap_true.
    EXIT.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Repeated package values -> one package (A06)
*& Consecutive lines with the same, non-empty Article No. and the same
*& Articles / Gross / Net values belong to one package: the repeated
*& values are cleared on the following lines. Same Article No. with
*& different values stays a new package (warning 068).
*&---------------------------------------------------------------------*
FORM collapse_repeated_groups CHANGING ct_lines    TYPE ty_upload_line_tab
                                       ct_messages TYPE bapiret2_t.

  DATA ls_prev    TYPE ty_upload_line.
  DATA lv_new_pkg TYPE abap_bool.

  LOOP AT ct_lines ASSIGNING FIELD-SYMBOL(<ls_line>).

    IF <ls_line>-artno IS NOT INITIAL AND <ls_line>-artno = ls_prev-artno.
      lv_new_pkg = abap_false.

      IF <ls_line>-art_filled = abap_true AND ls_prev-art_filled = abap_true.
        IF <ls_line>-art = ls_prev-art.
          <ls_line>-art_filled = abap_false.
          CLEAR <ls_line>-art.
        ELSE.
          PERFORM add_msg USING 'W' '068' <ls_line>-row <ls_line>-artno 'Articles' space
                          CHANGING ct_messages.
          lv_new_pkg = abap_true.
        ENDIF.
      ENDIF.

      IF <ls_line>-gwt_filled = abap_true AND <ls_line>-nwt_filled = abap_true
         AND ls_prev-gwt_filled = abap_true AND ls_prev-nwt_filled = abap_true.
        IF <ls_line>-gwt = ls_prev-gwt AND <ls_line>-nwt = ls_prev-nwt.
          <ls_line>-gwt_filled = abap_false.
          <ls_line>-nwt_filled = abap_false.
          CLEAR: <ls_line>-gwt, <ls_line>-nwt.
        ELSE.
          PERFORM add_msg USING 'W' '068' <ls_line>-row <ls_line>-artno 'Gross/Net Wt.' space
                          CHANGING ct_messages.
          lv_new_pkg = abap_true.
        ENDIF.
      ENDIF.

      " Compare the next line with the package values, not the cleared
      " ones - unless this line started a new package
      IF lv_new_pkg = abap_false.
        CONTINUE.
      ENDIF.
    ENDIF.

    ls_prev = <ls_line>.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel lines -> table rows with article / weight groups
*&---------------------------------------------------------------------*
FORM build_items USING    iv_packno   TYPE zsd_packno
                          it_lines    TYPE ty_upload_line_tab
                 CHANGING ct_items    TYPE ty_item_tab
                          ct_messages TYPE bapiret2_t.

  DATA lv_art_grp  TYPE i.
  DATA lv_wt_grp   TYPE i.
  DATA lv_pkg_type TYPE char20.
  DATA lv_prev_sn  TYPE zsd_sn.
  DATA lt_seen_sn  TYPE SORTED TABLE OF zsd_sn WITH UNIQUE KEY table_line.
  DATA ls_item     TYPE ty_item.

  CLEAR ct_items.

  LOOP AT it_lines INTO DATA(ls_line).

    " Sl.No. unique and ascending - groups depend on the row order
    INSERT ls_line-sn INTO TABLE lt_seen_sn.
    IF sy-subrc <> 0.
      PERFORM add_msg USING 'E' '021' ls_line-row ls_line-sn space space CHANGING ct_messages.
    ELSEIF ls_line-sn < lv_prev_sn.
      PERFORM add_msg USING 'E' '041' ls_line-row ls_line-sn space space CHANGING ct_messages.
    ENDIF.
    lv_prev_sn = ls_line-sn.

    IF ls_line-art_filled = abap_true.
      lv_art_grp = lv_art_grp + 1.
    ELSEIF lv_art_grp IS INITIAL.
      PERFORM add_msg USING 'E' '027' ls_line-row 'Articles' space space CHANGING ct_messages.
    ENDIF.

    " Gross and net always come together
    IF ls_line-gwt_filled = abap_true OR ls_line-nwt_filled = abap_true.
      IF ls_line-gwt_filled = abap_false OR ls_line-nwt_filled = abap_false.
        PERFORM add_msg USING 'E' '028' ls_line-row space space space CHANGING ct_messages.
      ENDIF.
      lv_wt_grp = lv_wt_grp + 1.
    ELSEIF lv_wt_grp IS INITIAL.
      PERFORM add_msg USING 'E' '027' ls_line-row 'Gross/Net Wt.' space space CHANGING ct_messages.
    ENDIF.

    " Package type is entered once per run of lines (A09)
    IF ls_line-pkg_type IS NOT INITIAL.
      lv_pkg_type = ls_line-pkg_type.
    ENDIF.

    CLEAR ls_item.
    ls_item-zsd_packno   = iv_packno.
    ls_item-zsd_sn       = ls_line-sn.
    ls_item-zsd_inv_sn   = ls_line-inv_sn.
    ls_item-zsd_artno    = ls_line-artno.
    ls_item-zsd_std      = ls_line-std.
    IF ls_line-art_filled = abap_true.
      ls_item-zsd_art    = ls_line-art.
    ENDIF.
    ls_item-zsd_art_grp  = lv_art_grp.
    ls_item-zsd_wt_grp   = lv_wt_grp.
    ls_item-zsd_pkg_type = lv_pkg_type.
    ls_item-zsd_partno   = ls_line-partno.
    ls_item-zsd_matdesc  = ls_line-matdesc.
    ls_item-zsd_qty      = ls_line-qty.
    IF ls_line-gwt_filled = abap_true.
      ls_item-zsd_gwt    = ls_line-gwt.
    ENDIF.
    IF ls_line-nwt_filled = abap_true.
      ls_item-zsd_nwt    = ls_line-nwt.
    ENDIF.
    ls_item-zsd_hsn      = ls_line-hsn.
    APPEND ls_item TO ct_items.

  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& FS 2.1: "Gross Weight is always equal to or greater than the Net
*& Weight ... Invalid Weight: Gross Weight cannot be less than Net
*& Weight." - checked per weight group (A07)
*&---------------------------------------------------------------------*
FORM check_weights USING    it_items    TYPE ty_item_tab
                   CHANGING ct_messages TYPE bapiret2_t.

  DATA lv_prev_grp TYPE i.
  DATA lt_detail   TYPE bapiret2_t.
  DATA lv_gross    TYPE string.
  DATA lv_net      TYPE string.
  DATA lv_sn       TYPE i.

  LOOP AT it_items INTO DATA(ls_item).
    " Weights are stored on the first line of each weight group only
    IF ls_item-zsd_wt_grp = lv_prev_grp.
      CONTINUE.
    ENDIF.
    lv_prev_grp = ls_item-zsd_wt_grp.

    IF ls_item-zsd_gwt < ls_item-zsd_nwt.
      PERFORM format_weight USING ls_item-zsd_gwt CHANGING lv_gross.
      PERFORM format_weight USING ls_item-zsd_nwt CHANGING lv_net.
      lv_sn = ls_item-zsd_sn.
      PERFORM add_msg USING 'E' '002' lv_sn lv_gross lv_net space CHANGING lt_detail.
    ENDIF.
  ENDLOOP.

  IF lt_detail IS NOT INITIAL.
    " The exact FS text first, then one line per failing group
    PERFORM add_msg USING 'E' '001' space space space space CHANGING ct_messages.
    APPEND LINES OF lt_detail TO ct_messages.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Quantity = STD x Articles - warning only (A08)
*&---------------------------------------------------------------------*
FORM check_quantities USING    it_items    TYPE ty_item_tab
                      CHANGING ct_messages TYPE bapiret2_t.

  DATA lv_articles TYPE zsd_art.
  DATA lv_prev_grp TYPE i.
  DATA lv_expected TYPE zsd_qty.
  DATA lv_qty      TYPE string.
  DATA lv_sn       TYPE i.
  DATA lv_std      TYPE i.
  DATA lv_art      TYPE i.

  LOOP AT it_items INTO DATA(ls_item).
    IF ls_item-zsd_art_grp <> lv_prev_grp.
      lv_articles = ls_item-zsd_art.
      lv_prev_grp = ls_item-zsd_art_grp.
    ENDIF.

    lv_expected = ls_item-zsd_std * lv_articles.
    IF lv_expected <> ls_item-zsd_qty.
      PERFORM format_quantity USING ls_item-zsd_qty CHANGING lv_qty.
      lv_sn  = ls_item-zsd_sn.
      lv_std = ls_item-zsd_std.
      lv_art = lv_articles.
      PERFORM add_msg USING 'W' '029' lv_sn lv_qty lv_std lv_art CHANGING ct_messages.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Totals of one packing list
*&---------------------------------------------------------------------*
FORM calc_totals USING    it_items  TYPE ty_item_tab
                 CHANGING cs_totals TYPE ty_totals.

  CLEAR cs_totals.

  LOOP AT it_items INTO DATA(ls_item).
    cs_totals-articles = cs_totals-articles + ls_item-zsd_art.
    cs_totals-quantity = cs_totals-quantity + ls_item-zsd_qty.
    cs_totals-gross    = cs_totals-gross    + ls_item-zsd_gwt.
    cs_totals-net      = cs_totals-net      + ls_item-zsd_nwt.

    READ TABLE cs_totals-pkg_types ASSIGNING FIELD-SYMBOL(<ls_pkg>)
         WITH KEY pkg_type = ls_item-zsd_pkg_type.
    IF sy-subrc <> 0.
      APPEND VALUE #( pkg_type = ls_item-zsd_pkg_type ) TO cs_totals-pkg_types ASSIGNING <ls_pkg>.
    ENDIF.
    <ls_pkg>-articles = <ls_pkg>-articles + ls_item-zsd_art.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& "Total 4074 Bundles" / "Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages"
*&---------------------------------------------------------------------*
FORM total_text USING    is_totals TYPE ty_totals
                CHANGING cv_text   TYPE string.

  DATA lv_type TYPE string.

  DATA(lt_types) = is_totals-pkg_types.
  DELETE lt_types WHERE pkg_type IS INITIAL AND articles = 0.

  CASE lines( lt_types ).
    WHEN 0.
      cv_text = |Total { is_totals-articles } Packages|.
    WHEN 1.
      lv_type = condense( CONV string( lt_types[ 1 ]-pkg_type ) ).
      IF lv_type IS INITIAL.
        lv_type = `Packages`.
      ENDIF.
      cv_text = |Total { is_totals-articles } { lv_type }|.
    WHEN OTHERS.
      cv_text = `Total`.
      LOOP AT lt_types INTO DATA(ls_type).
        lv_type = condense( CONV string( ls_type-pkg_type ) ).
        IF lv_type IS INITIAL.
          lv_type = `Packages`.
        ENDIF.
        IF sy-tabix > 1.
          cv_text = |{ cv_text } +|.
        ENDIF.
        cv_text = |{ cv_text } { ls_type-articles } { lv_type }|.
      ENDLOOP.
      cv_text = |{ cv_text } = { is_totals-articles } Packages|.
  ENDCASE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Items -> printed rows: HSN headings, items, package-type subtotals
*&---------------------------------------------------------------------*
FORM build_print_items USING    it_items    TYPE ty_item_tab
                                it_headings TYPE ty_hsn_heading_tab
                       CHANGING ct_print    TYPE zsd_tt_epack_af_itm.

  DATA lv_prev_hsn  TYPE zsd_hsn VALUE '#'.
  DATA lv_prev_art  TYPE i.
  DATA lv_prev_wt   TYPE i.
  DATA lv_run_type  TYPE char20.
  DATA lv_run_sum   TYPE i.
  DATA lv_subtotals TYPE abap_bool.
  DATA ls_totals    TYPE ty_totals.
  DATA ls_row       TYPE zsd_s_epack_af_itm.
  DATA ls_next      TYPE zsd_s_epack_af_itm.
  DATA lv_next      TYPE i.
  DATA lv_text      TYPE string.
  DATA lt_art_rows  TYPE HASHED TABLE OF ty_pkg_count WITH UNIQUE KEY grp.
  DATA lt_wt_rows   TYPE HASHED TABLE OF ty_pkg_count WITH UNIQUE KEY grp.
  DATA ls_count     TYPE ty_pkg_count.
  DATA lv_grp       TYPE i.

  CLEAR ct_print.

  DATA(lt_items) = it_items.
  SORT lt_items BY zsd_sn.

  " Subtotal rows only when more than one package type is used
  PERFORM calc_totals USING lt_items CHANGING ls_totals.
  lv_subtotals = xsdbool( lines( ls_totals-pkg_types ) > 1 ).

  LOOP AT lt_items INTO DATA(ls_item).

    IF lv_subtotals = abap_true AND sy-tabix > 1 AND ls_item-zsd_pkg_type <> lv_run_type.
      CLEAR ls_row.
      ls_row-row_type = gc_row_subtot.
      ls_row-articles = |{ lv_run_sum }|.
      ls_row-row_text = |<={ condense( CONV string( lv_run_type ) ) }|.
      ls_row-pkg_type = lv_run_type.
      APPEND ls_row TO ct_print.
      CLEAR lv_run_sum.
    ENDIF.

    IF ls_item-zsd_hsn <> lv_prev_hsn.
      CLEAR ls_row.
      ls_row-row_type = gc_row_heading.
      READ TABLE it_headings INTO DATA(ls_heading) WITH KEY hsn = ls_item-zsd_hsn.
      ls_row-row_text = COND #( WHEN sy-subrc = 0 THEN ls_heading-text
                                ELSE |HS CODE : { ls_item-zsd_hsn }| ).
      APPEND ls_row TO ct_print.
      lv_prev_hsn = ls_item-zsd_hsn.
    ENDIF.

    CLEAR ls_row.
    ls_row-row_type   = gc_row_item.
    ls_row-sl_no      = |{ CONV i( ls_item-zsd_sn ) }|.
    IF ls_item-zsd_inv_sn IS NOT INITIAL.
      ls_row-inv_sn = |{ CONV i( ls_item-zsd_inv_sn ) }|.
    ENDIF.
    ls_row-article_no = ls_item-zsd_artno.
    ls_row-std        = |{ CONV i( ls_item-zsd_std ) }|.
    ls_row-part_no    = ls_item-zsd_partno.
    ls_row-prod_desc  = ls_item-zsd_matdesc.
    PERFORM format_quantity USING ls_item-zsd_qty CHANGING lv_text.
    ls_row-quantity   = lv_text.
    ls_row-hsn        = ls_item-zsd_hsn.
    ls_row-pkg_type   = ls_item-zsd_pkg_type.
    ls_row-art_grp    = ls_item-zsd_art_grp.
    ls_row-wt_grp     = ls_item-zsd_wt_grp.
    ls_row-art_first  = xsdbool( ls_item-zsd_art_grp <> lv_prev_art ).
    ls_row-wt_first   = xsdbool( ls_item-zsd_wt_grp  <> lv_prev_wt ).
    " "Merged" cells: values on the first row of a group only
    IF ls_row-art_first = abap_true.
      ls_row-articles = |{ CONV i( ls_item-zsd_art ) }|.
    ENDIF.
    IF ls_row-wt_first = abap_true.
      PERFORM format_weight USING ls_item-zsd_gwt CHANGING lv_text.
      ls_row-gross_wt = lv_text.
      PERFORM format_weight USING ls_item-zsd_nwt CHANGING lv_text.
      ls_row-net_wt = lv_text.
    ENDIF.
    APPEND ls_row TO ct_print.

    " Rows per group (ART_ROWS / WT_ROWS)
    ls_count-grp = ls_item-zsd_art_grp.
    READ TABLE lt_art_rows ASSIGNING FIELD-SYMBOL(<ls_art>) WITH TABLE KEY grp = ls_count-grp.
    IF sy-subrc = 0.
      <ls_art>-rows = <ls_art>-rows + 1.
    ELSE.
      ls_count-rows = 1.
      INSERT ls_count INTO TABLE lt_art_rows.
    ENDIF.
    ls_count-grp = ls_item-zsd_wt_grp.
    READ TABLE lt_wt_rows ASSIGNING FIELD-SYMBOL(<ls_wt>) WITH TABLE KEY grp = ls_count-grp.
    IF sy-subrc = 0.
      <ls_wt>-rows = <ls_wt>-rows + 1.
    ELSE.
      ls_count-rows = 1.
      INSERT ls_count INTO TABLE lt_wt_rows.
    ENDIF.

    lv_prev_art = ls_item-zsd_art_grp.
    lv_prev_wt  = ls_item-zsd_wt_grp.
    lv_run_type = ls_item-zsd_pkg_type.
    lv_run_sum  = lv_run_sum + ls_item-zsd_art.

  ENDLOOP.

  IF lv_subtotals = abap_true AND lt_items IS NOT INITIAL.
    CLEAR ls_row.
    ls_row-row_type = gc_row_subtot.
    ls_row-articles = |{ lv_run_sum }|.
    ls_row-row_text = |<={ condense( CONV string( lv_run_type ) ) }|.
    ls_row-pkg_type = lv_run_type.
    APPEND ls_row TO ct_print.
  ENDIF.

  " Last row of each group (the form draws the bottom border there) and
  " the number of rows of the group
  LOOP AT ct_print ASSIGNING FIELD-SYMBOL(<ls_row>) WHERE row_type = gc_row_item.
    lv_next = sy-tabix + 1.
    CLEAR ls_next.
    READ TABLE ct_print INTO ls_next INDEX lv_next.
    <ls_row>-art_last = xsdbool( ls_next-row_type <> gc_row_item OR ls_next-art_grp <> <ls_row>-art_grp ).
    <ls_row>-wt_last  = xsdbool( ls_next-row_type <> gc_row_item OR ls_next-wt_grp  <> <ls_row>-wt_grp ).
    lv_grp = <ls_row>-art_grp.
    READ TABLE lt_art_rows INTO ls_count WITH TABLE KEY grp = lv_grp.
    <ls_row>-art_rows = COND #( WHEN sy-subrc = 0 THEN ls_count-rows ELSE 1 ).
    lv_grp = <ls_row>-wt_grp.
    READ TABLE lt_wt_rows INTO ls_count WITH TABLE KEY grp = lv_grp.
    <ls_row>-wt_rows = COND #( WHEN sy-subrc = 0 THEN ls_count-rows ELSE 1 ).
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Section heading of one HSN: the "Description of goods" line that
*& contains the HSN, otherwise "<T604N text> - HS CODE : <hsn>"
*&---------------------------------------------------------------------*
FORM heading_for_hsn USING    iv_hsn      TYPE csequence
                              it_goods    TYPE ty_text_tab
                              iv_hsn_text TYPE csequence
                     CHANGING cv_text     TYPE string.

  DATA lv_hsn_digits  TYPE string.
  DATA lv_line_digits TYPE string.
  DATA lv_hsn_fmt     TYPE zsd_hsn.

  PERFORM digits_only USING iv_hsn CHANGING lv_hsn_digits.

  IF lv_hsn_digits IS NOT INITIAL.
    LOOP AT it_goods INTO DATA(lv_line).
      PERFORM digits_only USING lv_line CHANGING lv_line_digits.
      IF lv_line_digits CS lv_hsn_digits.
        cv_text = lv_line.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDIF.

  PERFORM format_hsn USING iv_hsn CHANGING lv_hsn_fmt.
  IF iv_hsn_text IS INITIAL.
    cv_text = |HS CODE : { lv_hsn_fmt }|.
  ELSE.
    cv_text = |{ iv_hsn_text } - HS CODE : { lv_hsn_fmt }|.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& "39172390" -> "3917 2390" (format of the samples)
*&---------------------------------------------------------------------*
FORM format_hsn USING    iv_hsn TYPE csequence
                CHANGING cv_hsn TYPE zsd_hsn.

  DATA(lv_text) = replace( val = condense( CONV string( iv_hsn ) ) sub = ` ` with = `` occ = 0 ).

  IF strlen( lv_text ) = 8 AND lv_text CO '0123456789'.
    cv_hsn = |{ lv_text(4) } { lv_text+4(4) }|.
  ELSE.
    cv_hsn = condense( CONV string( iv_hsn ) ).
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
FORM format_weight USING    iv_weight TYPE zsd_gwt
                   CHANGING cv_text   TYPE string.
  cv_text = |{ iv_weight DECIMALS = 3 }|.
ENDFORM.

*&---------------------------------------------------------------------*
FORM format_quantity USING    iv_quantity TYPE zsd_qty
                     CHANGING cv_text     TYPE string.
  IF frac( iv_quantity ) = 0.
    cv_text = |{ CONV i( iv_quantity ) }|.
  ELSE.
    cv_text = |{ iv_quantity DECIMALS = 3 }|.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& DD-MM-YYYY, as on the samples
*&---------------------------------------------------------------------*
FORM format_date USING    iv_date TYPE d
                 CHANGING cv_text TYPE string.
  CLEAR cv_text.
  IF iv_date IS NOT INITIAL.
    cv_text = |{ iv_date+6(2) }-{ iv_date+4(2) }-{ iv_date(4) }|.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Multi-line Excel cell (Alt+Enter) -> lines
*&---------------------------------------------------------------------*
FORM split_lines USING    iv_text  TYPE csequence
                 CHANGING ct_lines TYPE ty_text_tab.

  DATA lt_raw TYPE ty_text_tab.

  CLEAR ct_lines.
  DATA(lv_text) = CONV string( iv_text ).
  DATA(lv_cr)   = substring( val = cl_abap_char_utilities=>cr_lf len = 1 ).
  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN lv_text WITH cl_abap_char_utilities=>newline.
  REPLACE ALL OCCURRENCES OF lv_cr IN lv_text WITH cl_abap_char_utilities=>newline.

  SPLIT lv_text AT cl_abap_char_utilities=>newline INTO TABLE lt_raw.

  LOOP AT lt_raw INTO DATA(lv_line).
    lv_line = condense( val = lv_line del = ` ` ).
    IF lv_line IS NOT INITIAL.
      APPEND lv_line TO ct_lines.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel cell -> number ("2107.2000000000003", "1,234.5")
*&---------------------------------------------------------------------*
FORM parse_number USING    iv_text  TYPE csequence
                  CHANGING cv_value TYPE decfloat34
                           cv_ok    TYPE abap_bool.

  CLEAR cv_value.
  cv_ok = abap_false.

  DATA(lv_text) = condense( val = CONV string( iv_text ) del = ` ` ).
  REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ``.

  IF lv_text IS INITIAL.
    cv_ok = abap_true.
    RETURN.
  ENDIF.

  TRY.
      cv_value = CONV decfloat34( lv_text ).
      cv_ok = abap_true.
    CATCH cx_sy_conversion_error.
      cv_ok = abap_false.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel cell -> date: Excel serial number (46275), DD.MM.YYYY,
*& DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD, YYYYMMDD
*&---------------------------------------------------------------------*
FORM parse_date USING    iv_text TYPE csequence
                CHANGING cv_date TYPE d
                         cv_ok   TYPE abap_bool.

  DATA lv_date   TYPE d.
  DATA lv_day    TYPE string.
  DATA lv_mon    TYPE string.
  DATA lv_year   TYPE string.
  DATA lv_serial TYPE decfloat34.
  DATA lv_num_ok TYPE abap_bool.

  CLEAR cv_date.
  cv_ok = abap_false.
  DATA(lv_text) = condense( val = CONV string( iv_text ) del = ` ` ).

  IF lv_text IS INITIAL.
    RETURN.
  ENDIF.

  IF strlen( lv_text ) = 8 AND lv_text CO '0123456789' AND ( lv_text(2) = '19' OR lv_text(2) = '20' ).
    lv_date = lv_text.                                    " YYYYMMDD
  ELSEIF lv_text CO '0123456789.'.
    " Excel serial date (base 30.12.1899)
    PERFORM parse_number USING lv_text CHANGING lv_serial lv_num_ok.
    IF lv_num_ok = abap_false OR lv_serial < 1 OR lv_serial > 2958465.
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
    cv_date = lv_date.
    cv_ok   = abap_true.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& "90000014, 90000015 & 90000020" -> invoice numbers (ALPHA format)
*&---------------------------------------------------------------------*
FORM split_invoices USING    iv_text    TYPE csequence
                    CHANGING ct_vbeln   TYPE ty_vbeln_tab
                             ct_invalid TYPE ty_text_tab.

  DATA lt_token TYPE ty_text_tab.

  CLEAR: ct_vbeln, ct_invalid.

  DATA(lv_text) = CONV string( iv_text ).
  REPLACE ALL OCCURRENCES OF REGEX `[,;&/\s]+` IN lv_text WITH ` `.
  SPLIT condense( lv_text ) AT ` ` INTO TABLE lt_token.

  LOOP AT lt_token INTO DATA(lv_token) WHERE table_line IS NOT INITIAL.
    IF lv_token CO '0123456789' AND strlen( lv_token ) <= 10.
      INSERT CONV vbeln_vf( |{ lv_token ALPHA = IN WIDTH = 10 }| ) INTO TABLE ct_vbeln.
    ELSE.
      APPEND lv_token TO ct_invalid.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
FORM digits_only USING    iv_text TYPE csequence
                 CHANGING cv_text TYPE string.
  cv_text = replace( val = iv_text regex = `[^0-9]` with = `` occ = 0 ).
ENDFORM.

*&---------------------------------------------------------------------*
*& "Buyer's Reference No. & Date:" -> "BUYERSREFERENCENODATE"
*&---------------------------------------------------------------------*
FORM normalize_heading USING    iv_text TYPE csequence
                       CHANGING cv_text TYPE string.
  cv_text = replace( val = to_upper( iv_text ) regex = `[^A-Z0-9]` with = `` occ = 0 ).
ENDFORM.

*&---------------------------------------------------------------------*
*& 1 -> A, 27 -> AA, 38 -> AL
*&---------------------------------------------------------------------*
FORM column_name USING    iv_col  TYPE i
                 CHANGING cv_name TYPE string.

  DATA lv_col TYPE i.
  DATA lv_rem TYPE i.

  CLEAR cv_name.
  lv_col = iv_col.
  WHILE lv_col > 0.
    lv_rem = ( lv_col - 1 ) MOD 26.
    cv_name = |{ sy-abcde+lv_rem(1) }{ cv_name }|.
    lv_col = ( lv_col - 1 - lv_rem ) DIV 26.
  ENDWHILE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Join non-empty parts with a separator
*&---------------------------------------------------------------------*
FORM join_parts USING    it_parts TYPE ty_text_tab
                         iv_sep   TYPE csequence
                CHANGING cv_text  TYPE string.

  CLEAR cv_text.
  LOOP AT it_parts INTO DATA(lv_part).
    lv_part = condense( lv_part ).
    IF lv_part IS INITIAL.
      CONTINUE.
    ENDIF.
    IF cv_text IS INITIAL.
      cv_text = lv_part.
    ELSE.
      cv_text = |{ cv_text }{ iv_sep }{ lv_part }|.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Multi-line text -> one STRING with line breaks (Adobe text field)
*&---------------------------------------------------------------------*
FORM text_block USING    iv_text TYPE csequence
                CHANGING cv_text TYPE string.

  DATA lt_lines TYPE ty_text_tab.

  PERFORM split_lines USING iv_text CHANGING lt_lines.
  cv_text = concat_lines_of( table = lt_lines sep = cl_abap_char_utilities=>newline ).

ENDFORM.

*&---------------------------------------------------------------------*
*& "No. of Pkgs." when not uploaded: "2342 Loose Pipes" / "1073 Boxes" /
*& "3415 Pkgs." from the package-type subtotals
*&---------------------------------------------------------------------*
FORM packages_text USING    is_totals TYPE ty_totals
                   CHANGING cv_text   TYPE string.

  DATA lt_lines TYPE ty_text_tab.
  DATA lv_type  TYPE string.

  LOOP AT is_totals-pkg_types INTO DATA(ls_type) WHERE pkg_type IS NOT INITIAL.
    lv_type = condense( CONV string( ls_type-pkg_type ) ).
    APPEND |{ ls_type-articles } { lv_type }| TO lt_lines.
  ENDLOOP.
  APPEND |{ is_totals-articles } Pkgs.| TO lt_lines.

  cv_text = concat_lines_of( table = lt_lines sep = cl_abap_char_utilities=>newline ).

ENDFORM.
