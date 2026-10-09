"! <p class="shorttext synchronized">Export Packing List - error</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Carries one T100 message (message class ZSD_EPACK) for the short text
"! and, optionally, the complete list of messages collected before the
"! error was raised (MT_MESSAGES) - the print program shows them all,
"! so the user can correct every problem in one go instead of one by one.
CLASS zcx_sd_epack DEFINITION
  PUBLIC
  INHERITING FROM cx_static_check
  CREATE PUBLIC.

  PUBLIC SECTION.

    INTERFACES if_t100_message.

    DATA mv_msgv1    TYPE symsgv READ-ONLY.
    DATA mv_msgv2    TYPE symsgv READ-ONLY.
    DATA mv_msgv3    TYPE symsgv READ-ONLY.
    DATA mv_msgv4    TYPE symsgv READ-ONLY.
    DATA mt_messages TYPE bapiret2_t READ-ONLY.

    METHODS constructor
      IMPORTING textid      LIKE if_t100_message=>t100key OPTIONAL
                previous    LIKE previous                 OPTIONAL
                mv_msgv1    TYPE symsgv                   OPTIONAL
                mv_msgv2    TYPE symsgv                   OPTIONAL
                mv_msgv3    TYPE symsgv                   OPTIONAL
                mv_msgv4    TYPE symsgv                   OPTIONAL
                mt_messages TYPE bapiret2_t               OPTIONAL.

    "! Raise message ZSD_EPACK/iv_msgno with up to four variables
    CLASS-METHODS raise
      IMPORTING iv_msgno    TYPE symsgno
                iv_v1       TYPE simple OPTIONAL
                iv_v2       TYPE simple OPTIONAL
                iv_v3       TYPE simple OPTIONAL
                iv_v4       TYPE simple OPTIONAL
                it_messages TYPE bapiret2_t OPTIONAL
      RAISING   zcx_sd_epack.

    "! Raise with the first error of a collected message list as text
    CLASS-METHODS raise_from_messages
      IMPORTING it_messages TYPE bapiret2_t
      RAISING   zcx_sd_epack.

ENDCLASS.


CLASS zcx_sd_epack IMPLEMENTATION.

  METHOD constructor.

    super->constructor( previous = previous ).

    me->mv_msgv1    = mv_msgv1.
    me->mv_msgv2    = mv_msgv2.
    me->mv_msgv3    = mv_msgv3.
    me->mv_msgv4    = mv_msgv4.
    me->mt_messages = mt_messages.

    CLEAR me->textid.
    if_t100_message~t100key = COND #( WHEN textid IS INITIAL
                                      THEN if_t100_message=>default_textid
                                      ELSE textid ).

  ENDMETHOD.


  METHOD raise.

    DATA(ls_key) = VALUE scx_t100key( msgid = zif_sd_epack=>gc_msgid
                                      msgno = iv_msgno
                                      attr1 = 'MV_MSGV1'
                                      attr2 = 'MV_MSGV2'
                                      attr3 = 'MV_MSGV3'
                                      attr4 = 'MV_MSGV4' ).

    RAISE EXCEPTION TYPE zcx_sd_epack
      EXPORTING textid      = ls_key
                mv_msgv1    = |{ iv_v1 }|
                mv_msgv2    = |{ iv_v2 }|
                mv_msgv3    = |{ iv_v3 }|
                mv_msgv4    = |{ iv_v4 }|
                mt_messages = it_messages.

  ENDMETHOD.


  METHOD raise_from_messages.

    " Show the first error as the exception text; the caller displays
    " the complete list from MT_MESSAGES.
    LOOP AT it_messages INTO DATA(ls_message) WHERE type CA 'EAX'.
      EXIT.
    ENDLOOP.
    IF ls_message IS INITIAL.
      ls_message = VALUE #( it_messages[ 1 ] OPTIONAL ).
    ENDIF.

    DATA(ls_key) = VALUE scx_t100key( msgid = ls_message-id
                                      msgno = ls_message-number
                                      attr1 = 'MV_MSGV1'
                                      attr2 = 'MV_MSGV2'
                                      attr3 = 'MV_MSGV3'
                                      attr4 = 'MV_MSGV4' ).

    RAISE EXCEPTION TYPE zcx_sd_epack
      EXPORTING textid      = ls_key
                mv_msgv1    = ls_message-message_v1
                mv_msgv2    = ls_message-message_v2
                mv_msgv3    = ls_message-message_v3
                mv_msgv4    = ls_message-message_v4
                mt_messages = it_messages.

  ENDMETHOD.

ENDCLASS.
