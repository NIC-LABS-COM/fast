
*&---------------------------------------------------------------------*
*& Report Z_GET_ALL_PACKAGES
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT Z_GET_ALL_PACKAGES.

TYPES: BEGIN OF ty_package,
         devclass TYPE tdevc-devclass,
         ctext    TYPE tdevct-ctext,
       END OF ty_package.

DATA: lt_packages TYPE STANDARD TABLE OF ty_package,
      ls_package  TYPE ty_package,
      lt_output   TYPE STANDARD TABLE OF string,
      lv_fullpath TYPE string,
      lv_line     TYPE string.

START-OF-SELECTION.

  SELECT
    package~devclass AS devclass,
    text~ctext       AS ctext
    FROM tdevc AS package
    LEFT OUTER JOIN tdevct AS text
      ON package~devclass = text~devclass
     AND text~spras       = @sy-langu
    WHERE package~devclass LIKE 'Z%'
       OR package~devclass LIKE 'Y%'
       OR package~devclass = 'LOCAL'
       OR package~devclass = '$TMP'
    ORDER BY package~devclass
    INTO TABLE @lt_packages.

  IF sy-subrc <> 0 OR lt_packages IS INITIAL.
    MESSAGE 'Nenhum package encontrado.' TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  LOOP AT lt_packages INTO ls_package.
    CLEAR lv_line.

    IF ls_package-ctext IS INITIAL.
      lv_line = ls_package-devclass.
    ELSE.
      lv_line = ls_package-devclass.
    ENDIF.

    APPEND lv_line TO lt_output.
  ENDLOOP.

  lv_fullpath = 'C:\temp\packages_tdevc.txt'.

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
    MESSAGE |{ lines( lt_output ) } packages baixados em { lv_fullpath }| TYPE 'S'.
  ELSE.
    MESSAGE 'Erro ao fazer download do arquivo.' TYPE 'E'.
  ENDIF.Z
