*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_SEL
*&---------------------------------------------------------------------*
*& Selection screen and selection-screen events
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_xls   RADIOBUTTON GROUP src DEFAULT 'X' USER-COMMAND src.
PARAMETERS p_file  TYPE rlgrap-filename LOWER CASE MODIF ID xls.
PARAMETERS p_saved RADIOBUTTON GROUP src.
SELECT-OPTIONS s_vbeln  FOR gv_vbeln  MODIF ID sav.
SELECT-OPTIONS s_packno FOR gv_packno MODIF ID sav.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-b02.
PARAMETERS p_fauto RADIOBUTTON GROUP fmt DEFAULT 'X'.
PARAMETERS p_fwith RADIOBUTTON GROUP fmt.
PARAMETERS p_fwout RADIOBUTTON GROUP fmt.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-b03.
PARAMETERS p_prev  RADIOBUTTON GROUP out DEFAULT 'X' USER-COMMAND out.
PARAMETERS p_print RADIOBUTTON GROUP out.
PARAMETERS p_pdf   RADIOBUTTON GROUP out.
PARAMETERS p_dest  TYPE rspopname.
PARAMETERS p_copy  TYPE tdcopies DEFAULT 1.
PARAMETERS p_dir   TYPE rlgrap-filename LOWER CASE MODIF ID pdf.
PARAMETERS p_check AS CHECKBOX.
SELECTION-SCREEN END OF BLOCK b3.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
*----------------------------------------------------------------------*
  LOOP AT SCREEN.
    CASE screen-group1.
      WHEN 'XLS'.
        screen-active = COND #( WHEN p_xls = abap_true THEN '1' ELSE '0' ).
      WHEN 'SAV'.
        screen-active = COND #( WHEN p_saved = abap_true THEN '1' ELSE '0' ).
      WHEN 'PDF'.
        screen-active = COND #( WHEN p_pdf = abap_true THEN '1' ELSE '0' ).
      WHEN OTHERS.
        CONTINUE.
    ENDCASE.
    MODIFY SCREEN.
  ENDLOOP.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM f4_file CHANGING p_file.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_dir.
*----------------------------------------------------------------------*
  PERFORM f4_folder CHANGING p_dir.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN.
*----------------------------------------------------------------------*
  IF sy-ucomm = 'ONLI' OR sy-ucomm = 'PRIN'.
    IF p_xls = abap_true AND p_file IS INITIAL.
      MESSAGE e017(zsd_epack) WITH '-' 'Excel file'.
    ENDIF.
    IF p_saved = abap_true AND s_vbeln[] IS INITIAL AND s_packno[] IS INITIAL.
      MESSAGE e003(zsd_epack).
    ENDIF.
    IF p_pdf = abap_true AND p_check = abap_false AND p_dir IS INITIAL.
      MESSAGE e054(zsd_epack).
    ENDIF.
  ENDIF.
