*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_SEL
*&---------------------------------------------------------------------*
*& Selection screen: the Excel file only. Output is a print preview of
*& the Adobe Form on device LP01 (printable from the preview).
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_file TYPE rlgrap-filename LOWER CASE OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM f4_file CHANGING p_file.
