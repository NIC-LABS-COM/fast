*&---------------------------------------------------------------------*
*& Report Z_BUSCA_REPORTS
*&---------------------------------------------------------------------*
REPORT z_busca_reports.

TYPES: BEGIN OF ty_report,
         obj_name TYPE tadir-obj_name,
       END OF ty_report.

DATA: lt_reports  TYPE STANDARD TABLE OF ty_report,
      ls_report   TYPE ty_report,
      lt_output   TYPE STANDARD TABLE OF string,
      lv_fullpath TYPE string,
      lv_line     TYPE string.

START-OF-SELECTION.
A
  SELECT obj_name
    FROM tadir
    INTO TABLE @lt_reports
    WHERE pgmid    = 'R3TR'
      AND object   = 'PROG'
      AND ( obj_name LIKE 'Z%' OR obj_name LIKE 'Y%' )
    ORDER BY obj_name.

  IF sy-subrc <> 0 OR lt_reports IS INITIAL.
    MESSAGE 'Nenhum report Z ou Y encontrado na TADIR.' TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  LOOP AT lt_reports INTO ls_report.
    CLEAR lv_line.
    lv_line = ls_report-obj_name.
    APPEND lv_line TO lt_output.
  ENDLOOP.

  lv_fullpath = 'C:\temp\reports.txt'.

  CALL METHOD cl_gui_frontend_services=>gui_download
    EXPORTING
      filename                = lv_fullpath
      filetype                = 'ASC'
    CHANGING
      data_tab                = lt_output
    EXCEPTIONS
      file_write_error        = 1
      no_batch                = 2
      gui_refuse_filetransfer = 3
      invalid_type            = 4
      no_authority            = 5
      unknown_error           = 6
      header_not_allowed      = 7
      separator_not_allowed   = 8
      filesize_not_allowed    = 9
      header_too_long         = 10
      dp_error_create         = 11
      dp_error_send           = 12
      dp_error_write          = 13
      unknown_dp_error        = 14
      access_denied           = 15
      dp_out_of_memory        = 16
      disk_full               = 17
      dp_timeout              = 18
      file_not_found          = 19
      dataprovider_exception  = 20
      control_flush_error     = 21
      OTHERS                  = 22.

  IF sy-subrc = 0.
    MESSAGE |{ lines( lt_output ) } reports Z/Y baixados em { lv_fullpath }| TYPE 'S'.
  ELSE.
    MESSAGE 'Erro ao fazer download do arquivo.' TYPE 'E'.
  ENDIF.
