  METHOD get_data.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - receives internal table from report
    "
    " UPDATES
    "   - 08.09.2025 :: new way of distributing data
    "   - 23.09.2025 :: set customers into new table for kunnr
    "   - 23.09.2025 :: set subprojects into new table for projects
    "--------------------------------------------------------------------------------------------

    CREATE OBJECT o_xl.

    get_print_setup( ).

    IF lv_flag_honorarios EQ 'X'.
      LOOP AT gt_dados INTO DATA(ls_honorarios) WHERE lstar EQ 'A001'.
        MOVE-CORRESPONDING ls_honorarios TO gs_data.
        APPEND gs_data TO gt_honorarios.
      ENDLOOP.
      SORT gt_honorarios BY pep_mask ASCENDING.
    ELSEIF lv_flag_despesas EQ 'X'.
      LOOP AT gt_dados INTO DATA(ls_despesas) WHERE lstar NE 'A001'.
        MOVE-CORRESPONDING ls_despesas TO gs_data.
        APPEND gs_data TO gt_despesas.
      ENDLOOP.
    ELSE.
      LOOP AT gt_dados INTO DATA(ls_honorarios2) WHERE lstar EQ 'A001'.
        MOVE-CORRESPONDING ls_honorarios2 TO gs_data.
        APPEND gs_data TO gt_honorarios.
      ENDLOOP.
      SORT gt_honorarios BY pep_mask ASCENDING.
      LOOP AT gt_dados INTO DATA(ls_despesas2) WHERE lstar NE 'A001'.
        MOVE-CORRESPONDING ls_despesas2 TO gs_data.
        APPEND gs_data TO gt_despesas.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.
