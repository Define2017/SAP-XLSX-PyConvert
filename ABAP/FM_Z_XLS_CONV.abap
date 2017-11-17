FUNCTION z_xls_conv.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     REFERENCE(IV_SRC_FOLDER) TYPE  CHAR0064
*"     REFERENCE(IV_LOG_FOLDER) TYPE  CHAR0064
*"     REFERENCE(IV_DST_FOLDER) TYPE  CHAR0064
*"----------------------------------------------------------------------
* This function module will execute the SM69 external command Z_XLS_CONV which will call the python script xlsx_convert.py
* Build our local working variables
  DATA: lv_status   TYPE extcmdexex-status,
        iv_sm69_cmd TYPE sxpgcolist-parameters,
        lv_retcode  TYPE extcmdexex-exitcode,
        lt_output   TYPE STANDARD TABLE OF btcxpm.

  FIELD-SYMBOLS: <lfs_output> LIKE LINE OF lt_output.
  "Execute the selected command:

  IF ( iv_src_folder IS NOT INITIAL ) AND ( iv_log_folder IS NOT INITIAL ) AND ( iv_dst_folder IS NOT INITIAL ).
***** We build the command line string

    CONCATENATE '/usr/sap/SAPBW/intf/in/xlsx_convert.py -l ''' iv_log_folder ''' -s ''' iv_src_folder ''' -d ''' iv_dst_folder ''''  INTO iv_sm69_cmd.
    
    REPLACE ALL OCCURRENCES OF '''' IN iv_sm69_cmd WITH ' '.
    
***** This section of code will get populate the temporary table t_move with the output of a LS command in the directory
    CALL FUNCTION 'SXPG_COMMAND_EXECUTE'
      EXPORTING
        commandname                   = 'Z_XLS_CONV'
        additional_parameters         = iv_sm69_cmd
        operatingsystem               = 'Linux'
        stdout                        = 'X'
        stderr                        = 'X'
        terminationwait               = 'X'
      IMPORTING
        status                        = lv_status
        exitcode                      = lv_retcode
      TABLES
        exec_protocol                 = lt_output
      EXCEPTIONS
        no_permission                 = 1
        command_not_found             = 2
        parameters_too_long           = 3
        security_risk                 = 4
        wrong_check_call_interface    = 5
        program_start_error           = 6
        program_termination_error     = 7
        x_error                       = 8
        parameter_expected            = 9
        too_many_parameters           = 10
        illegal_command               = 11
        wrong_asynchronous_parameters = 12
        cant_enq_tbtco_entry          = 13
        jobcount_generation_error     = 14
        OTHERS                        = 15.

  ENDIF.

  IF sy-subrc NE 0.
    " Display the error message
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno " this is display the system message automatically when ever exp are hit
    WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ELSE.
    WRITE: 'Succesfully executed python script. Check xls_conv.log for more details.'.
  ENDIF.



ENDFUNCTION.