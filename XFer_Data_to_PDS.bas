Attribute VB_Name = "XFer_Data_to_PDS"
Option Explicit
Public first_test_ref As Integer, last_test_ref As Integer



Sub Transfer_Data()

' The purpose of the this module is to read limit data and insert into a pds
' file that has been converted to excel format.

' Note this module has the following constraints  -
'         Assumes the max limit files in a project is 25
'         Assumes that the test name or test description column is named "TESTNAME"
'         Assumes that the FT, room temp, low limit  column is named  "R LTL"
'         Assumes that the FT, room temp, high limit column is named  "R UTL"
'         Assumes that the FT, cold temp, low limit  column is named  "C LTL"
'         Assumes that the FT, cold temp, high limit column is named  "C UTL"
'         Assumes that the FT, hot  temp, low limit  column is named  "H LTL"
'         Assumes that the FT, hot  temp, high limit column is named  "H UTL"
'         Assumes that the QC, room temp, low limit  column is named  "QC LTL"
'         Assumes that the QC, room temp, high limit column is named  "QC UTL"
'         Assumes that the QC, hot  temp, low limit  column is named  "QCH LTL"
'         Assumes that the QC, hot  temp, high limit column is named  "QCH UTL"
'         Assumes that the QC, cold temp, low limit  column is named  "QCC LTL"
'         Assumes that the QC, cold temp, high limit column is named  "QCC UTL"
'         Assumes that the limit format column title is named "DFORMAT"
'         Assumes that the Data Sheet Variable Map is spaced no more than 200 lines
'         below the top of excel version of the pds (not overly excessive
'         Preferences or Binning blocks)
'         * Add software to check actual position of the Variable Map
'         Assumes no more than 20 columns in a pds file
'         Furthermore, the routine assumes the presence of "XFer Lmts" column available
'         to the user to omit limits from the transfer process.

   Dim proceed As Integer

   MsgBox "This utility requires the following pds file column naming!" & vbNewLine & "1) TESTNAME" & vbNewLine & "2) R/C/H LTL" & vbNewLine & "3) R/C/H UTL" & vbNewLine & "4) DFORMAT", vbCritical + vbExclamation, "PDS Column Format"
   MsgBox "This utility requires the following QC column naming!" & vbNewLine & "1) QCR/C/H LTL" & vbNewLine & "2) QCR/C/H UTL" & vbNewLine, vbCritical + vbExclamation, "PDS Column Format"
   ' If the user acknowledges that "Yes" they have applied proper formatting, a 6 is returned.
   '        Button  Constant    Value
   '          OK      vbOK        1
   '        Cancel  vbCancel      2
   '         Abort   vbAbort      3
   '         Retry   vbRetry      4
   '        Ignore   vbIgnore     5
   '          Yes     vbYes       6
   '           No     vbNo        7
   proceed = MsgBox("Has the proper pds file column formating been applied to all pds files?", vbYesNo + vbCritical + vbExclamation, "PDS Column Format")
   
  'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
  
  ' Get the number of pds files converted to excel format. Call get_num_pds_files function
  ' to return the number of pds files in the project and the work sheet names of each pds
  ' file
  
   Dim num_pds_files As Integer
   Dim pds_names(1 To 25) As String  ' Code is limited to projects with 25 pds files or less
  
   num_pds_files = get_num_pds_files(pds_names)
    
   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
   
    ' Check to make sure the limits file spreadsheet has been read and copied to the "Test_Limits"
    ' of this workbook.
   
    Dim num_worksheets As Integer
    Dim test_limits_exist As Boolean
    Dim sheet_check_index As Integer
    Dim sheet_name As String
   
    num_worksheets = ThisWorkbook.Worksheets.count
    
    test_limits_exist = False
   
    For sheet_check_index = 1 To num_worksheets
            sheet_name = ThisWorkbook.Sheets(sheet_check_index).Name
            
            ' If not the PDS Utilities sheet or the Test Limits sheet, check cell A1 to verify a pds file sheet
            ' For pds files, cell A1 should be "[Datasheet Preferences]"
            If (StrComp(sheet_name, "Test_Limits", vbTextCompare) = 0) Then
                test_limits_exist = True
           End If
    Next sheet_check_index
   
   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
   
   
   ' Read the Datasheet Variable Map block of each pds file and store the column reference
   ' of the Test Name, LTL, UTL, QC LTL, QC UTL and if applicable, QC Hot LTL, QC Hot UTL, QC Cold LTL, QC Cold UTL
   'Dim test_name_col As Integer
   'Dim r_ltl_col, r_utl_col, c_ltl_col, c_utl_col, h_ltl_col, h_utl_col As Integer
   'Dim qcr_ltl_col, qcr_utl_col, qcc_ltl_col, qcc_utl_col, qch_ltl_col, qch_utl_col As Integer
   'Dim dformat_col As Integer
   Dim xfer_col As Integer
   Dim var_map_ndx As Integer
   Dim temp_worksheet_name As String
   Dim limit_file_ndx As Integer
   'Dim pds_column_data(1 To 9) As Integer
   'Dim pds_column_data(1 To 10) As Integer
   Dim pds_column_data(1 To 14) As Integer
   Dim test_count, pds_index As Integer
   
   Dim temp_pds_test_name As String
   Dim edited_temp_pds_test_name As String
   Dim temp_limit_test_name As String
   Dim string_match As Boolean
   Dim prod_names As String
   Dim InStr_Rslt As Integer
   Dim temp_values As String
   Dim precision_format As Single
   Dim test_limits_wbk As Workbook
   
   'Set test_limits_wbk = Workbooks.Open("C:\ets\project_limits\PDS_UtilitiesRevC.xlsm", True, True)
   
   'test_limits_wbk = ThisWorkbook.Worksheets("Test_Limits")
    
   Dim mod_columns As Integer
   
   Dim limit_columns As Collection
   
        'test_name_col = -1
        'r_ltl_col = -1
        'r_utl_col = -1
        'c_ltl_col = -1
        'c_utl_col = -1
        'h_ltl_col = -1
        'h_utl_col = -1
        'qcr_ltl_col = -1
        'qcr_utl_col = -1
        'qcc_ltl_col = -1
        'qcc_utl_col = -1
        'qch_ltl_col = -1
        'qch_utl_col = -1
        
        
        ' For each pds file determine the column location in order to determine where to write data
        ' Do a check to ensure the pds file columns (limits) match the limit spreadhseet columns
        '                  ** Advise the user of mismatch
        ' Transfer limits from limit spreadsheet to pds files.
        '                  ** ALL limits in a particular category (FT, QC..) spreadsheet should have a
        '                     a corresponding entry in the pds.  Warn the user if this is not the case.
        ' pds_column_data ->
        '
        '           pds_column_data(1) = "Test Name"
        '           pds_column_data(2) = "R LTL"
        '           pds_column_data(3) = "R UTL"
        '           pds_column_data(4) = "C LTL"
        '           pds_column_data(5) = "C UTL"
        '           pds_column_data(6) = "H LTL"
        '           pds_column_data(7) = "H UTL"
        '           pds_column_data(8) = "QCR LTL"
        '           pds_column_data(9) = "QCR UTL"
        '           pds_column_data(10) = "QCC LTL"
        '           pds_column_data(11) = "QCC UTL"
        '           pds_column_data(12) = "QCH LTL"
        '           pds_column_data(13) = "QCH UTL"
        
        '           pds_column_data(14) = "DFORMAT"
        
        'temp_worksheet = ThisWorkbook.temp_sheet
        
        Dim pds_count As Integer
        Dim tests_updated, tests_no_match As Integer
        Dim test_updated_this_loop As Boolean
        Dim pds_column_message As String
        Dim pds_column_check As Integer
        
          
        ' Don't bother trying to transfer limits if user has acknowledged imptoprt proper pds column formating
        If (proceed = 6) Then
          ' Don't bother trying to transfer limits if the production limits were not read.
          If (test_limits_exist = True) Then
              For pds_count = 1 To num_pds_files
        
                test_updated_this_loop = False
                tests_updated = 0
                tests_no_match = 0
        
                ' Select a pds file to edit
                temp_worksheet_name = pds_names(pds_count)
                
                ' Locate the column data for transfering llimits. If the pds columns have been improperly named
                ' the routine will return a 0. Otherwise it will return the number of columns that may be modified
                ' if corresponding limits are found on the test limits worksheet.
                mod_columns = get_pds_column_data(pds_column_data, temp_worksheet_name)
                    
                
                ' Advise the user of the limits that will be transferred to each PDS file based on the availability
                ' 1) Limits type (R LTL,R UTL, QCH LTL, etc)in the PE Limits file AND 2) corresponding data column
                ' in the PDS file.
                pds_column_message = ""
                
                If (pds_column_data(2) > 0) Then pds_column_message = "1)R LTL"
                If (pds_column_data(3) > 0) Then pds_column_message = pds_column_message + ", R UTL" & vbCr
                If (pds_column_data(4) > 0) Then pds_column_message = pds_column_message + "2) C LTL"
                If (pds_column_data(5) > 0) Then pds_column_message = pds_column_message + ", C UTL" & vbCr
                If (pds_column_data(6) > 0) Then pds_column_message = pds_column_message + "3) H LTL"
                If (pds_column_data(7) > 0) Then pds_column_message = pds_column_message + ", H UTL" & vbCr
                If (pds_column_data(8) > 0) Then pds_column_message = pds_column_message + "4) QCR LTL"
                If (pds_column_data(9) > 0) Then pds_column_message = pds_column_message + ", QCR UTL" & vbCr
                If (pds_column_data(10) > 0) Then pds_column_message = pds_column_message + "5) QCC LTL"
                If (pds_column_data(11) > 0) Then pds_column_message = pds_column_message + ", QCC UTL" & vbCr
                If (pds_column_data(12) > 0) Then pds_column_message = pds_column_message + "6) QCH LTL"
                If (pds_column_data(13) > 0) Then pds_column_message = pds_column_message + ", QCH UTL" & vbCr
                
                  
                MsgBox "Limit file:" + temp_worksheet_name & vbCr + "The following limits will be transfered" & vbCr + pds_column_message, vbCritical + vbOKOnly, "PDS File Transfer"
                
                 
                If (mod_columns > 0) Then
                
                   ' If the Production limits worksheet is not read, the count variable is not set and a loop
                   ' below will not be executed. Address here for now. Exit the loop and cease executing any code
                   ' transfering limits to pds files.
                   If count = 0 Then
                        MsgBox "Production Limits have not been read. Please Reset the project and re=read the production limits! ", vbCritical, "Production Limits Work Sheet Not Read"
                        Exit For
                   End If
                   
                    ' find the the number of tests in the pds file
                    test_count = get_pds_test_count(temp_worksheet_name, pds_column_data(1))
            
                    ' With column references known and worksheet name (pds_names(num_pds_sheets)) known
                    ' Should now be able transfer limits. Maybe add a routine to do string compare (InStr)
                    ' For now, all limits copied into the pds should be doubles
                
                    ' Below are the required steps
                    ' 1) Search pds for last row entry of a pds limit OR test name using the pds_column provided
                    ' 2) Loop from 1 to least number of elements in either the pds or limit file durign the loop -
                    '                     - do a string compare of the test name when a match is found copy limits to pds
                
                    ' Nested loop selecting each test on the pds then indexing throught the limit worksheet for a test name match in order to transfer limits
                    Dim xfer_limit As String
                    Dim prod_lim_names_cycle As Integer
                    Dim duplicate_pds_tstname_count As Integer
                    duplicate_pds_tstname_count = Check_4Duplicate_Test_Names(temp_worksheet_name, first_test_ref, last_test_ref, pds_column_data(1))
                  
                    For pds_index = first_test_ref To last_test_ref
                 
                        test_updated_this_loop = False
                        temp_pds_test_name = UCase(ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(1)))    'Convert PDS Test name to uppercase
                        
                       
                        precision_format = ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(14))
                        
                        edited_temp_pds_test_name = Replace(temp_pds_test_name, Chr(34), "")
                   
                        For prod_lim_names_cycle = 2 To count + 1
                        
                            InStr_Rslt = 1    ' Initialize to 1 and modify based on conditions defined below.
                        
                            prod_names = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, 2).Value
                            
                            If (edited_temp_pds_test_name = "") Then
                               InStr_Rslt = 0
                            Else
                               ' If the test limit test name(prod_names) is contained in the pds test name(edited_temp_pds_test_name) then InStr_Rslt > 0   -> write limits
                                InStr_Rslt = InStr(1, prod_names, edited_temp_pds_test_name, vbTextCompare)
                                
                                 ' Read Xfer Limit column of the Test_limits worksheet on per test basis to determine if limits should be transferred to pds files
                                 ' If the user has designated the test to not have limits imported to the pds, over-write the InStr_Rslt variable=0
                                xfer_limit = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, 3).Value
                                If xfer_limit = "No" Then
                                  InStr_Rslt = 0
                                End If
                            End If
                            
                            If InStr_Rslt > 0 Then
                            
                                ' R LTL limits must be present in the PE Limits file to proceed
                                If r_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "D").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, room temp, lower limit column (r_ltl_col) AND the R LTL column exists (pds_column_data(2) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to R LTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(2) > 0)) Then
                                    
                                        temp_values = Set_precision(precision_format, temp_values)  ' dont set precision unless a number
                                    
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(2)) = temp_values
                                        ' Keep track of the number of tests with updated limits by test name (not per specific limits)
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                       
                                    Else ' Warn the user of a limit that is not a number whether an empty cell or random keyboard symbol
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The room LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' R UTL limits must be present in the PE Limits file to proceed
                                If r_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "E").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, room temp, upper limit column (r_utl_col) AND the R UTL column exists (pds_column_data(3) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to R LTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(3) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(3)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                        
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The room UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                                
                                    
                                ' C LTL limits must be present in the PE Limits file to proceed
                                If c_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "F").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, cold temp, lower limit column (c_ltl_col) AND the C LTL column exists (pds_column_data(4) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to C LTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(4) > 0)) Then
                                    
                                        temp_values = Set_precision(precision_format, temp_values)  ' dont set precision unless a number
                                    
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(4)) = temp_values
                                        ' Keep track of the number of tests with updated limits by test name (not per specific limits)
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                       
                                    Else ' Warn the user of a limit that is not a number whether an empty cell or random keyboard symbol
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The cold LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' C UTL limits must be present in the PE Limits file to proceed
                                If c_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "G").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, cold temp, upper limit column (c_ltl_col) AND the C UTL column exists (pds_column_data(5) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to C UTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(5) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(5)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                        
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The colf UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                                
                                
                                ' H LTL limits must be present in the PE Limits file to proceed
                                If h_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "H").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, hot temp, lower limit column (h_ltl_col) AND the H LTL column exists (pds_column_data(6) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to H LTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(3) > 6)) Then
                                    
                                        temp_values = Set_precision(precision_format, temp_values)  ' dont set precision unless a number
                                    
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(6)) = temp_values
                                        ' Keep track of the number of tests with updated limits by test name (not per specific limits)
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                       
                                    Else ' Warn the user of a limit that is not a number whether an empty cell or random keyboard symbol
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The hot LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' H UTL limits must be present in the PE Limits file to proceed
                                If h_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "I").Value
                                    
                                    ' If the Test_Limits worksheet contains FT, hot temp, upper limit column (h_utl_col) AND the H UTL column exists (pds_column_data(7) > 0) in
                                    ' the PDS file and the user has indicated that they want the test limits transferred, copy limits to H UTL column of selected pds file
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(7) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(7)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                        
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The hot UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                                                             
                                
                                ' QCR LTL limits must be present in the PE Limits file to proceed
                                If qcr_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "J").Value
                                       
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(8) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(8)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The QCR_LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' QCR UTL limits must be present in the PE Limits file to proceed
                                If qcr_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "K").Value
                                    
    
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(9) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(9)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The QCR_UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' QCC LTL limits must be present in the PE Limits file to proceed
                                If qcc_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "L").Value
                                    

                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(10) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(10)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                     Else
                                        If (Not (temp_pds_test_name = "")) Then
                                           MsgBox "The QCC_LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                ' QCC UTL limits must be present in the PE Limits file to proceed
                                If qcc_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "M").Value
                                    
                            
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(11) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(11)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                            MsgBox "The QCC_UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                             
                                ' QCH LTL limits must be present in the PE Limits file to proceed
                                If qch_ltl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "N").Value
                                    
                        
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(12) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(12)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                    Else
                                        If (Not (temp_pds_test_name = "")) Then
                                            MsgBox "The QCH_LTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                        End If
                                    End If
                                End If
                            
                                 ' QCH UTL limits must be present in the PE Limits file to proceed
                                If qch_utl_col > 0 Then
                                    temp_values = ThisWorkbook.Worksheets("Test_Limits").Cells(prod_lim_names_cycle, "O").Value
                                    
                        
                                    If ((IsNumeric(temp_values) = True) And (xfer_limit = "Yes") And (pds_column_data(13) > 0)) Then
                                        temp_values = Set_precision(precision_format, temp_values)
                                        ThisWorkbook.Worksheets(temp_worksheet_name).Cells(pds_index, pds_column_data(13)) = temp_values
                                        If (Not (test_updated_this_loop)) Then
                                           tests_updated = tests_updated + 1
                                           test_updated_this_loop = True
                                        End If
                                
                                     Else
                                        If (Not (temp_pds_test_name = "")) Then
                                            MsgBox "The QCH_UTL Production Limit for ''" & edited_temp_pds_test_name & "'' is formatted incorrectly on the production limits sheet. The PDS value will remain unchanged.", vbCritical, "Production Limit Error"
                                         End If
                                     End If
                                 End If
                        
                                 Exit For
                              Else
                            End If
                           
                        Next prod_lim_names_cycle
                    
                        If InStr_Rslt <= 0 Then
                           tests_no_match = tests_no_match + 1
                           'MsgBox "There was no match found for ''" & edited_temp_pds_test_name & "'' on the production limits sheet.", vbCritical, "Test Name Not Found"
                        End If
                    Next pds_index
                ElseIf (mod_columns = 0) Then
                    MsgBox "PDs file -> ''" & temp_worksheet_name & "'' column titles not formatted correctly!", vbCritical, "PDS File Format Error"
                    MsgBox "Please use TESTNAME, DFORMAT, R LTL, R UTL, C LTL, C UTL, H LTL, H UTL, QCR LTL, QCR UTL, QCH LTL, QCH UTL, QCC LTL and/or QCC UTL!", vbCritical, "PDS File Format Error"
                End If
             Next pds_count
          End If
        End If
End Sub


Public Function get_index(worksheet_name As String, col_name As String) As Integer

 '  The purpose of this routine is to determine the column location index for data defined
 ' in the pds file. Because the user has the freedom to define columns as they wish,
 ' no assumptions can be made relative to column location of a particular parameter
 ' (Below is TYPICAL pds column definition.)
 ' This is accomplished by first locating [Datasheet Variable Map] text
 ' in column A of a converted pds file. Then a text search is done to find the position
 ' of the column of interest which are defined in the rows of data in the Variable Map.
 ' Exam -> A search of the definitions below would identify the "HiLim" or upper final
 ' test limit in column 5. Therefore column 5 would be updated with FT upper limits provided
 ' by Product Engineering
 
        '[Datasheet Variable Map]
        '"integer","TestNmbr","Test Number"
        '"integer","SubTestNmbr","Subtest Number"
        '"character","DLogDesc","Datalog Description"
        '"double","R LTL","RM Low Limit"
        '"double","R UTL","RM High Limit"
        '"double","QCR LTL","QCR Low Limit"
        '"double","QCR UTL","QCR High Limit"
        '"double","var3","setup variable 3"
        '"double","DFormat","Data Format"
        '"character","Units","Units"
        '"character","LoFBin","Low Fail Bin"
        '"character","HiFBin","High Fail Bin"
        
  Dim var_map As String
  Dim read_pds_cell As String
  Dim read_ds_var_map_cell As String
  Dim ds_var_map_index As Integer
  Dim var_map_index As Integer
  Dim column_num As Integer
  Dim loop_exit As Boolean
  Dim string_comp As Integer
  
 
    column_num = -1
    loop_exit = False
  
    ' Search the first 200 cells in column A to find the Data Sheet Variable Map block
    For ds_var_map_index = 1 To 200
  
         read_pds_cell = ThisWorkbook.Worksheets(worksheet_name).Cells(ds_var_map_index, 1)
         ' If the Data Sheet Variable Map Block is found look an additional 20 cells for the
         ' the variable of interest
         If read_pds_cell = "[Datasheet Variable Map]" Then
         
              For var_map_index = 1 To 20
                
                  read_ds_var_map_cell = ThisWorkbook.Worksheets(worksheet_name).Cells(ds_var_map_index + var_map_index, 2)
                  string_comp = InStr(1, read_ds_var_map_cell, col_name, vbTextCompare)
                  If (string_comp > 0) Then
                       column_num = var_map_index
                       loop_exit = True
                       Exit For
                  End If
                  If var_map_index = 20 Then  ' If we've found the Variable Map but not the data column then exit
                    loop_exit = True
                    column_num = 0
                   End If
                  
              Next var_map_index
         End If
         If loop_exit = True Then Exit For
         
    Next ds_var_map_index
  
  get_index = column_num    ' return column index
  
End Function


Public Function get_num_pds_files(ByRef pds_names() As String)

 ' Determine how many pds limit files are in the project by searching tabs.
 ' Read Cell A1. If Cell A1 = "[Datasheet Preferences]" then the worksheet is a
 ' pds limit file.
 ' If a pds limit file, save the worksheet name in a string array.
  
   Dim num_sheets, num_pds_sheets, num_limit_files As Integer
   Dim sheet_index As Integer
   Dim pds_check As String
   Dim sheet_name As String
   
   ThisWorkbook.Activate    ' set current workbook as active
                                       
   num_sheets = ThisWorkbook.Worksheets.count
   num_pds_sheets = 0
   
   For sheet_index = 1 To num_sheets
            sheet_name = ThisWorkbook.Sheets(sheet_index).Name
            ' If not the PDS Utilities sheet or the Test Limits sheet, check cell A1 to verify a pds file sheet
            ' For pds files, cell A1 should be "[Datasheet Preferences]"
            If ((StrComp(sheet_name, "PDS Utilities", vbTextCompare) <> 0) And (StrComp(sheet_name, "Test_Limits", vbTextCompare) <> 0)) Then
                pds_check = ThisWorkbook.Worksheets(sheet_index).Cells(1, 1)
                If pds_check = "[Datasheet Preferences]" Then
                    num_pds_sheets = num_pds_sheets + 1
                    pds_names(num_pds_sheets) = sheet_name
                End If
           End If
   Next sheet_index
   
   get_num_pds_files = num_pds_sheets


End Function


' This function receives the name of the pds file and determines the column location of the limitss
' Column data is stored in the pds_column_data array
Public Function get_pds_column_data(ByRef pds_column_data() As Integer, pds_file_name As String)

 Dim limit_file_ndx As Integer
 Dim array_index1 As Integer
 Dim array_upper_bound As Integer
 Dim temp_worksheet As String
 Dim array_integer1 As Integer
 
 
       ' Init array so all values are -1
          array_upper_bound = UBound(pds_column_data)
          
          For array_integer1 = 1 To array_upper_bound
              pds_column_data(array_integer1) = 0
          Next array_integer1
          
          
          ' Capture column location information. If no corresponding column, the function returns a 0
          ' Increment by to compensate for "" in first cell of the excel converted pds
          ' Do not increment is a 0 is returned
          pds_column_data(1) = get_index(pds_file_name, "TESTNAME")  ' .cpp and pds must reference/contain Test_Name
          pds_column_data(1) = IIf(pds_column_data(1) > 0, pds_column_data(1) + 1, pds_column_data(1))
          ' ALL pds files MUST contain a "Test_Name" column!
          If pds_column_data(1) = 0 Then
              MsgBox pds_file_name & " - Critical Error -> No Test_Name column found!", vbCritical, "PDS Column Format"
          End If
          
          
          pds_column_data(2) = get_index(pds_file_name, "R LTL")        ' .cpp and pds must reference lower room temp FT limit as LTL
          pds_column_data(2) = IIf(pds_column_data(2) > 0, pds_column_data(2) + 1, pds_column_data(2))
          
          ' Omit for now
          'If pds_column_data(2) = 0 Then
          '    MsgBox pds_file_name & "  No final test, room temp, low test limit(R LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(3) = get_index(pds_file_name, "R UTL")        ' .cpp and pds must reference upper room temp FT limit as UTL
          pds_column_data(3) = IIf(pds_column_data(3) > 0, pds_column_data(3) + 1, pds_column_data(3))
          
          ' Omit for now
          'If pds_column_data(3) = 0 Then
          '    MsgBox pds_file_name & "  No final test,room temp, high test limit(R UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(4) = get_index(pds_file_name, "C LTL")        ' .cpp and pds must reference lower room temp FT limit as LTL
          pds_column_data(4) = IIf(pds_column_data(4) > 0, pds_column_data(4) + 1, pds_column_data(4))
          
          ' Omit for now
          'If pds_column_data(4) = 0 Then
          '    MsgBox pds_file_name & "  No final test,cold temp, low test limit(LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(5) = get_index(pds_file_name, "C UTL")        ' .cpp and pds must reference upper room temp FT limit as UTL
          pds_column_data(5) = IIf(pds_column_data(5) > 0, pds_column_data(5) + 1, pds_column_data(5))
          
          ' Omit for now
          'If pds_column_data(5) = 0 Then
          '    MsgBox pds_file_name & "  No final test,cold temp, high test limit(C UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(6) = get_index(pds_file_name, "H LTL")        ' .cpp and pds must reference lower room temp FT limit as LTL
          pds_column_data(6) = IIf(pds_column_data(4) > 0, pds_column_data(6) + 1, pds_column_data(6))
          
          ' Omit for now
          'If pds_column_data(6) = 0 Then
          '    MsgBox pds_file_name & "  No final test,high temp, low test limit(H LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(7) = get_index(pds_file_name, "H UTL")        ' .cpp and pds must reference upper room temp FT limit as UTL
          pds_column_data(7) = IIf(pds_column_data(7) > 0, pds_column_data(7) + 1, pds_column_data(7))
          
          ' Omit for now
          'If pds_column_data(7) = 0 Then
          '    MsgBox pds_file_name & "  No final test,high temp, high test limit(H UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          
          pds_column_data(8) = get_index(pds_file_name, "QCR LTL")    ' .cpp and pds must reference lower Room-QC limit as QCR_LTL
          pds_column_data(8) = IIf(pds_column_data(8) > 0, pds_column_data(8) + 1, pds_column_data(8))
          
          ' Omit for now
          'If pds_column_data(8) = 0 Then
          '    MsgBox pds_file_name & "  No QC test, room temp, lower test limit(QCR LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(9) = get_index(pds_file_name, "QCR UTL")    ' .cpp and pds must reference upper Room-QC limit as QCR_UTL
          pds_column_data(9) = IIf(pds_column_data(9) > 0, pds_column_data(9) + 1, pds_column_data(9))
          
          ' Omit for now
          'If pds_column_data(9) = 0 Then
          '    MsgBox pds_file_name & " No QC test, room temp, upper test limit(QCR UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(10) = get_index(pds_file_name, "QCC LTL")    ' .cpp and pds must reference lower Cold-QC limit as QCC_LTL
          pds_column_data(10) = IIf(pds_column_data(10) > 0, pds_column_data(10) + 1, pds_column_data(10))
          
          ' Omit for now
          'If pds_column_data(10) = 0 Then
          '    MsgBox pds_file_name & " No QC test, cold temp, lower test limit(QCC LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(11) = get_index(pds_file_name, "QCC UTL")    ' .cpp and pds must reference upper Cold-QC limit as QCC_UTL
          pds_column_data(11) = IIf(pds_column_data(11) > 0, pds_column_data(11) + 1, pds_column_data(11))
          
          ' Omit for now
          'If pds_column_data(11) = 0 Then
          '    MsgBox pds_file_name & " No QC test, cold temp, upper test limit(QCC UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(12) = get_index(pds_file_name, "QCH LTL")    ' .cpp and pds must reference lower Hot-QC limit as QCH_LTL
          pds_column_data(12) = IIf(pds_column_data(12) > 0, pds_column_data(12) + 1, pds_column_data(12))
          
          ' Omit for now
          'If pds_column_data(12) = 0 Then
          '    MsgBox pds_file_name & " No QC test, hot temp, lower test limit(QCH_LTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
          pds_column_data(13) = get_index(pds_file_name, "QCH UTL")    ' .cpp and pds must reference upper Hot-QC limit as QCH_UTL
          pds_column_data(13) = IIf(pds_column_data(13) > 0, pds_column_data(13) + 1, pds_column_data(13))
          
          ' Omit for now
          'If pds_column_data(13) = 0 Then
          '    MsgBox pds_file_name & " No QC test, hot temp, upper test limit(QCH_UTL) column found!", vbOKOnly, "PDS Column Format"
          'End If
          
        
          pds_column_data(14) = get_index(pds_file_name, "DFORMAT")   ' .cpp and pds must reference/contain DFORMAT
          pds_column_data(14) = IIf(pds_column_data(14) > 0, pds_column_data(14) + 1, pds_column_data(14))
          
          ' ALL pds files MUST contain a "DFORMAT" column!
          If pds_column_data(14) = 0 Then
              MsgBox pds_file_name & " - Critical Error -> No DFORMAT column found!", vbCritical, "PDS Column Format"
          End If
          
          get_pds_column_data = 0
          
          'Return the the total number of data columns that MAY be modified
          For array_integer1 = 2 To (array_upper_bound - 1) ' Check from upper bound -1 -> Do not include DFORFAMT as a column to be modified
              If pds_column_data(array_integer1) > 0 Then
                 get_pds_column_data = get_pds_column_data + 1
              End If
          Next array_integer1
          
          ' If the Test Name column header or DFormat columns are not labeled correctly then there is no way
          ' to transfer limits or properly format limits. Return 0 indicating no modifiable data columns until corrected.
          If (pds_column_data(1) = 0) Then
               get_pds_column_data = 0
          ElseIf (pds_column_data(14) = 0) Then
               get_pds_column_data = 0
              MsgBox pds_file_name & " Not updated due to column naming error!", vbCritical, "PDS Column Format"
          End If
                 
End Function


' Find the number of tests in pds file by looking for a test name between the first defined
' test function and the end of last defined test function
Private Function get_pds_test_count(temp_sheet As String, test_name_column As Integer)
'Private Function get_pds_test_count(temp_sheet As String, test_name_column As Integer, first_test_index As Integer, last_test_index As Integer)


    Dim search_index
    Dim first_test_index, last_test_index, test_count As Integer
    Dim temp_worksheet As Worksheet
    Dim temp_cell As String
    Dim first, last As Integer
    Dim str_compare As Variant
    
    Dim comp_string As String
    
    
       ThisWorkbook.Activate
       
        comp_string = "Function="
         
       'temp_worksheet = ThisWorkbook.Worksheets.temp_sheet
    
       For first_test_index = 1 To 1000
       
          temp_cell = ThisWorkbook.Worksheets(temp_sheet).Cells(first_test_index, 1)
       
          'temp_cell = temp_worksheet.Cells(first_test_index, 1)
          
          str_compare = InStr(1, temp_cell, comp_string, vbTextCompare)
          
          If ((str_compare > 0) And (Not IsNull(str_compare))) Then
              'If (Not IsNull(str_compare)) Then
          
                  first = first_test_index + 1
                  first_test_ref = first
              'End If

              Exit For
          End If
       Next first_test_index
       
       comp_string = "[TEST Order]"
       
       For last_test_index = first To 3000            ' The Test Order block should be before pds cell row 3000
       
         temp_cell = ThisWorkbook.Worksheets(temp_sheet).Cells(last_test_index, 1)
       
          'temp_cell = temp_worksheet.Cells(last_test_index, 1)
        str_compare = InStr(1, temp_cell, comp_string, vbTextCompare)
          
        If ((str_compare > 0) And (Not IsNull(str_compare))) Then
       
          'If (InStr(temp_cell, "[TEST Order]", vbTestCompare) > 0 And InStr(temp_cell, "[TEST Order]", vbTestCompare) <> Null) Then
          
              last = last_test_index - 1
              last_test_ref = last
              Exit For
          End If
       Next last_test_index
       
       
       get_pds_test_count = 0
    
       For test_count = first To last
       
           temp_cell = ThisWorkbook.Worksheets(temp_sheet).Cells(test_count, test_name_column)
        
           If (ThisWorkbook.Sheets(temp_sheet).Cells(test_count, test_name_column) <> "") Then
               get_pds_test_count = get_pds_test_count + 1
            End If
        Next test_count
               
End Function

Public Function Set_precision(precision As Single, passed_value As String)

  Dim temp_string As String
  Dim prec_string As String
  Dim result As Integer
  Dim prec As Integer
     
     If IsNumeric(passed_value) Then
            prec_string = CStr(precision)
            prec = Len(prec_string)
            If prec = 1 Then
                prec_string = "0"
            ElseIf prec > 1 Then
                prec_string = Right(prec_string, 1)
            End If
            'result = InStr(prec_string, "4", vbTextCompare)
            'If (result > 0) Then
            '   result = Right(prec_string, 1)
            'End If
            
            prec = CInt(prec_string)
            
            temp_string = FormatNumber(passed_value, prec, vbFalse, vbFalse, vbFalse)
            
            'for a temp_string value less than 1 w/ precision of 0 (not numbers to right of decimal point) set to 0
            ' to avoid a conversion error when converting the string to a single.
            If (temp_string = "") Then
              temp_string = "0"    ' manage the case where a
            End If
            
            Set_precision = CSng(temp_string)
     End If


End Function
