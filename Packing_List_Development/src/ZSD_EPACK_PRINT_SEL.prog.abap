*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_SEL
*&---------------------------------------------------------------------*
*& Selection screen and selection-screen events
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_file  TYPE rlgrap-filename LOWER CASE.
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
    IF p_file IS INITIAL.
      MESSAGE e017(zsd_epack) WITH '-' 'Excel file'.
    ENDIF.
    IF p_pdf = abap_true AND p_check = abap_false AND p_dir IS INITIAL.
      MESSAGE e054(zsd_epack).
    ENDIF.
  ENDIF.
