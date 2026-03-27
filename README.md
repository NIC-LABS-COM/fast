*&---------------------------------------------------------------------*
*& Report Z_BUSCA_REQUESTS
*&---------------------------------------------------------------------*
REPORT z_busca_requests.

TYPES: BEGIN OF ty_request,
         request     TYPE e070-trkorr,
         description TYPE e07t-as4text,
       END OF ty_request.

DATA: lt_requests TYPE STANDARD TABLE OF ty_request,
      ls_request  TYPE ty_request,
      lt_output   TYPE STANDARD TABLE OF string,
      lv_fullpath TYPE string,
      lv_line     TYPE string.

START-OF-SELECTION.

  SELECT
    request~trkorr      AS request,
    description~as4text AS description
    FROM e070 AS request
    LEFT OUTER JOIN e07t AS description
      ON request~trkorr = description~trkorr
    WHERE request~trfunction = 'K'
      AND request~trstatus   = 'D'
    ORDER BY request~trkorr
    INTO TABLE @lt_requests.

  IF sy-subrc <> 0 OR lt_requests IS INITIAL.
    MESSAGE 'Nenhuma request encontrada.' TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  LOOP AT lt_requests INTO ls_request.
    CLEAR lv_line.

    IF ls_request-description IS INITIAL.
      lv_line = ls_request-request.
    ELSE.
      lv_line = |{ ls_request-request } ; { ls_request-description }|.
    ENDIF.

    APPEND lv_line TO lt_output.
  ENDLOOP.

  lv_fullpath = 'C:\temp\workbench_requests.txt'.

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
    MESSAGE |{ lines( lt_output ) } requests baixadas em { lv_fullpath }| TYPE 'S'.
  ELSE.
    MESSAGE 'Erro ao fazer download do arquivo.' TYPE 'E'.
  ENDIF.
