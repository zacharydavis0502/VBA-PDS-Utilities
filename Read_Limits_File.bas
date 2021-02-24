Attribute VB_Name = "Read_Limits_File"
Option Explicit
Public count As Integer
Public r_ltl_col, r_utl_col, c_ltl_col, c_utl_col, h_ltl_col, h_utl_col, qcr_ltl_col, qcr_utl_col, qcc_ltl_col, qcc_utl_col, qch_ltl_col, qch_utl_col, units_col As Integer
Public testnum_col, testname_col As Integer


Sub Read_Limits()

    'Open the project directory and read the limit
    'file provided by the Product engineer. The file name MUST be
    ' renamed "PRODUCT_LIMITS"
    ' Before we open the project directory, verify that the directory folder
    ' exists. Then we'll verify the limits file and at least one limit file (.pds)
    ' file exists in the directory.

    ' Check that the path exists
    Dim Path As String
    Dim Folder_name As String
    Dim response As VbMsgBoxResult
    Dim folder_exists As Boolean
    
    ' Always start with the Prod Limits workbook closed
    
      If Is_File_Open("C:\ets\project_limits\project_pds_limits.xlsm") Then
         Workbooks("C:\ets\project_limits\project_pds_limits.xlsm").Close SaveChanges:=False
      ElseIf Is_File_Open("C:\ets\project_limits\project_pds_limits.xlsx") Then
         Workbooks("C:\ets\project_limits\project_pds_limits.xlsx").Close SaveChanges:=False
    End If
    
    Path = ("C:\ets\project_limits")
    Folder_name = Dir(Path, vbDirectory)
    If Folder_name = vbNullString Then
        response = MsgBox("C:\ets\project_limits does not exist. Would you like to create it?", vbYesNo, "Create Path?")
        Select Case response
            Case vbYes
                VBA.FileSystem.MkDir (Path)
                folder_exists = True
             Case Else
                folder_exists = False
        End Select
    Else
        MsgBox ("Folder C:\ets\project_limits exists")
        folder_exists = True
    End If
    
          
    ' Check the folder exists
    Dim filename As String
    Dim direct_filename As String

    If (folder_exists) Then

            If VBA.FileSystem.Dir("C:\ets\project_limits\project_pds_limits.xlsm") <> VBA.Constants.vbNullString Then
                filename = "C:\ets\project_limits\project_pds_limits.xlsm"
                direct_filename = "project_pds_limits.xlsm"
            ElseIf VBA.FileSystem.Dir("C:\ets\project_limits\project_pds_limits.xlsx") <> VBA.Constants.vbNullString Then
                filename = "C:\ets\project_limits\project_pds_limits.xlsx"
                direct_filename = "project_pds_limits.xlsx"
                'MsgBox ("project_pds_limits file does not exist in C:\ets\project_limits")
            Else
                MsgBox "project_pds_limits file does not exist in C:\ets\project_limits", vbCritical, "Production Limits Error"
            End If
            
            If filename <> VBA.Constants.vbNullString Then
                MsgBox ("project_pds_limits file found")
                Dim array_index As Integer
                Dim cold_index As Integer
                Dim hot_index As Integer
                Dim src_data As Workbook
                Dim testnum, testname, temp_name As String
                Dim r_ltl, r_utl, c_ltl, c_utl, h_ltl, h_utl  As String
                Dim qcr_ltl, qcr_utl, qcc_ltl, qcc_utl, qch_ltl, qch_utl, units As String
                
                ' Define search strings that will be used to reference relevant data columns in the Production Limits file
                testnum = "Test Id"
                testname = "Test Name"
                r_ltl = "Room LTL"
                r_utl = "Room UTL"
                c_ltl = "Cold LTL"
                c_utl = "Cold UTL"
                h_ltl = "Hot LTL"
                h_utl = "Hot UTL"
                
                qcr_ltl = "Room QC LTL"
                qcr_utl = "Room QC UTL"
                qcc_ltl = "Cold QC LTL"
                qcc_utl = "Cold QC UTL"
                qch_ltl = "Hot QC LTL"
                qch_utl = "Hot QC UTL"
                units = "Units"
        
                ' Initialize column refernces as null strings
                testnum_col = 0
                testname_col = 0
                r_ltl_col = 0
                r_utl_col = 0
                c_ltl_col = 0
                c_utl_col = 0
                h_ltl_col = 0
                h_utl_col = 0
                qcr_ltl_col = 0
                qcr_utl_col = 0
                qcc_ltl_col = 0
                qcc_utl_col = 0
                qcc_utl_col = 0
                qch_ltl_col = 0
                qch_utl_col = 0
                units_col = 0
                
                ' check if C:\ets\project_limits\project_pds_limits.xlsm is already open.
                ' If so do not re-open or else you get an error
                
                
                'define and open the limit file spread sheet for data transfer for read only access
                Set src_data = Workbooks.Open(filename, True, True)
                                
                'Verify that a "Production Limits" worksheet exists in the project_pds_limits workbook
                Dim prod_lim_sheet As String
                Dim hot_lim_sheet As String
                Dim cold_lim_sheet As String
                
                Dim prod_limits_exist As Boolean
                Dim hot_limits_exist As Boolean
                Dim cold_limits_exist As Boolean
                
                prod_lim_sheet = "Production Limits"
                hot_lim_sheet = "Hot Limits"
                cold_lim_sheet = "Cold Limits"
                
                prod_limits_exist = False
                hot_limits_exist = False
                cold_limits_exist = False
                ' init prod_limits_exist as false
                prod_limits_exist = worksheet_exists(prod_lim_sheet)
                hot_limits_exist = worksheet_exists(hot_lim_sheet)
                cold_limits_exist = worksheet_exists(cold_lim_sheet)
                
                
                ' Add check/throw up a msg box to make sure a Production Limits worksheets is contained in the file
                ' Respond if prod_limits_exist is false
                
                ThisWorkbook.Activate  ' set current workbook as active so data gets copied back from the limits file
                                       ' to the user project work book (workbook being used to update pds limits)
                   
                If (prod_limits_exist Or hot_limits_exist Or cold_limits_exist) Then
                          
                    ' Index across row #1 for column headings defined above. Limit search from column A to Z
                    ' If not a tri-temp solution the ht and cld qc limit string will be  ""
                     If (prod_limits_exist) Then
                        For array_index = 1 To 26                                   ' Row  column
                            temp_name = src_data.Worksheets(prod_lim_sheet).Cells(1, array_index).Value
                    
                            ' store column references for columns of interest
                            If temp_name = testnum Then testnum_col = array_index
                            If temp_name = testname Then testname_col = array_index
                            If temp_name = units Then units_col = array_index
                            If temp_name = r_ltl Then r_ltl_col = array_index
                            If temp_name = r_utl Then r_utl_col = array_index
                            If temp_name = qcr_ltl Then qcr_ltl_col = array_index
                            If temp_name = qcr_utl Then qcr_utl_col = array_index
                        
                        Next array_index
                    End If
                    
                    If (cold_limits_exist) Then
                        For cold_index = 1 To 26
                            temp_name = src_data.Worksheets(cold_lim_sheet).Cells(1, cold_index).Value
                        
                            If temp_name = testnum Then testnum_col = cold_index
                            If temp_name = testname Then testname_col = cold_index
                            If temp_name = units Then units_col = cold_index
                            If temp_name = c_ltl Then c_ltl_col = cold_index
                            If temp_name = c_utl Then c_utl_col = cold_index
                            If temp_name = qcc_ltl Then qcc_ltl_col = cold_index
                            If temp_name = qcc_utl Then qcc_utl_col = cold_index
                    
                        Next cold_index
                    End If
                        
                    If (hot_limits_exist) Then
                        For hot_index = 1 To 26
                            temp_name = src_data.Worksheets(hot_lim_sheet).Cells(1, hot_index).Value
                        
                            If temp_name = testnum Then testnum_col = hot_index
                            If temp_name = testname Then testname_col = hot_index
                            If temp_name = units Then units_col = hot_index
                            If temp_name = h_ltl Then h_ltl_col = hot_index
                            If temp_name = h_utl Then h_utl_col = hot_index
                            If temp_name = qch_ltl Then qch_ltl_col = hot_index
                            If temp_name = qch_utl Then qch_utl_col = hot_index
                    
                        Next hot_index
                    End If
                
                    ' Add a worksheet to transfer the limit data to, so limits are stored
                    ' as a matter of record and future reference. The added worksheet will
                    ' be named Test_Limits to be referenced by code transfering limit data
                    ' The added worksheet will be appended to the end of other worksheets -
                    ' do not delete/remove other worksheets the user may want/need.
                
                    Dim sheet As Worksheet
                    For Each sheet In ActiveWorkbook.Worksheets
                        If sheet.Name = "Test_Limits" Then
                            Application.DisplayAlerts = False
                            Worksheets("Test_Limits").Delete
                            Application.DisplayAlerts = True
                        End If
                    Next sheet
                    
                    Dim num_sheets As Integer
                
                    num_sheets = ThisWorkbook.Worksheets.count
                    Sheets.Add After:=Sheets(Sheets.count)
                    Sheets(num_sheets + 1).Select
                    Sheets(num_sheets + 1).Name = "Test_Limits"
                
                    ' Check # of limits and transfer to current workbook here
                    'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                
                    src_data.Activate
                
               
                    Dim i, exit_count As Integer
                    Dim names As String
                
                    count = 0
                    exit_count = 0

                    ' Search for up to 2002 unique test names (tests). Breaks in the limit file are allowed
                    ' for up to 10 rows. After 10 rows, if there is no test name is found, the test number
                    ' counting process is terminated and the test count at the point the last name was read
                    ' is reported as the total number of tests w/ limits
                    Dim prod_names_alt As String
                    
                    For i = 2 To 2002
                        If (prod_limits_exist) Then
                            names = src_data.Worksheets(prod_lim_sheet).Cells(i, testname_col).Value
                            prod_names_alt = prod_lim_sheet
                        ElseIf (hot_limits_exist) Then
                            names = src_data.Worksheets(hot_lim_sheet).Cells(i, testname_col).Value
                            prod_names_alt = hot_lim_sheet
                        ElseIf (cold_limits_exist) Then
                            names = src_data.Worksheets(cold_lim_sheet).Cells(i, testname_col).Value
                            prod_names_alt = cold_lim_sheet
                        End If
                        
                        
                        ' Search Test Name column for test names to determine number of tests in the
                        ' for which there are limits
                        If (Not names = "") Then
                            count = count + 1
                        ElseIf (names = "") Then
                            exit_count = exit_count + 1
                        End If
                    
                        If exit_count = 10 Then
                            Exit For
                        End If
                    
                    Next i
                
                    MsgBox "Test Limits - " & count & " tests.", vbInformation, "Prod. Eng. Test Limits"
                   
                    'NUM_PROD_LIMITS = count
                
                    Dim current_text As String
                    Dim current_limit As Double
                    Dim paste_workbook As Workbook
                    Dim copycount As Integer
                    
                    Set paste_workbook = ThisWorkbook
                      
                    'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                
                    'COPY TEST ID'S'
                    If testnum_col > 0 Then
                        For copycount = 1 To count + 1
                            current_text = src_data.Worksheets(prod_names_alt).Cells(copycount, testnum_col).Value
                            If (IsEmpty(current_text) = False) Then      'Only update the pds if the cell in limits file is not empty
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "A").Value = current_text
                            End If
                    
                        Next copycount
                    End If
                
                    'copy_limits_to_utilities(src_column_index As Integer, dest_column_index As Integer, num_limits As Integer)
                
                    'Dim mytest As Boolean
               
                    'If testnum_col > 0 Then
                    '        mytest = copy_limits_to_utilities(testnum_col, 1, count)
                    'End If
                    
                    
                    'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                
                    'COPY TEST_NAMES'
                    If testname_col > 0 Then
                        For copycount = 1 To count + 1
                            current_text = src_data.Worksheets(prod_names_alt).Cells(copycount, testname_col).Value
                             If (IsEmpty(current_text) = False) Then
                                 paste_workbook.Worksheets("Test_Limits").Cells(copycount, "B").Value = current_text
                                 
                                 'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                                 
                                 ' While transferring Test names from the Product Engineering Limits file to the PDS_Utilities worksheet, create a column
                                 ' that can be used by the user to manually omit the transfer of limits for certain/specific tests. The code below provides
                                 ' a validation list, in a column adjacent to the test name with "Yes" and "No" options.
                                 ' This status of column will be queeried during the transfer of limits to define which tests will have limits transfered
                                
                                 If (copycount = 1) Then
                                     paste_workbook.Worksheets("Test_Limits").Cells(copycount, 3).Value = "XFer"
                                 End If
                                 If (copycount > 1) Then
                                   paste_workbook.Worksheets("Test_Limits").Cells(copycount, 3).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
                                   paste_workbook.Worksheets("Test_Limits").Cells(copycount, 3).Value = "Yes"
                                 End If
                                
                                'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                                       
                             End If
                    
                        Next copycount
                     End If
                
                    'If testname_col > 0 Then
                    '    mytest = copy_limits_to_utilities(testname_col, 2, NUM_PROD_LIMITS)
                    'End If
                    
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                 
                   'Copy Room Temp, Final Test Lower Limit to the "Test Limits" worksheet
                    Dim formated_string As String
                    If r_ltl_col > 0 Then
                       For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Production Limits").Cells(copycount, r_ltl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "D").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                                    ' If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "D").Value = current_text
                                ' In certain instance, the formating of the limits from product engineering may contain
                                ' too many digits past the decimal point. So manage formating here.
                                ' Format the numbers in the RT_LowLim column after writing the column header,
                                ' once the first (number) limit is encountered
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("D2:D2002").Select
                                    Range("D2:D2002").NumberFormat = "0.000000000"
                                End If
                                 
                           End If
                       Next copycount
                    End If
                   
                   'If r_ltl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                   'Copy Room Temp, Final Test Upper Limit to the "Test Limits" worksheet
                   If r_utl_col > 0 Then
                      For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Production Limits").Cells(copycount, r_utl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "E").Value = current_text
                                
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "E").Value = current_text
                               
                                If (IsNumeric(current_text)) Then
                                    If (copycount = 2) Then
                                        paste_workbook.Worksheets("Test_Limits").Activate
                                        ' Set formatting for first 2000 limits
                                        Range("E2:E2002").Select
                                        Range("E2:E2002").NumberFormat = "0.000000000"
                                    End If
                                End If
                                
                           End If
                      Next copycount
                      
                   End If
                   
                   'If r_utl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                 
                   'Copy Cold Temp, Final Test Lower Limit to the "Test Limits" worksheet
                    If c_ltl_col > 0 Then
                       For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Cold Limits").Cells(copycount, c_ltl_col).Value
                           
                          If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "F").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "F").Value = current_text
                                ' In certain instance, the formating of the limits from product engineering may contain
                                ' too many digits past the decimal point. So manage formating here.
                                ' Format the numbers in the RT_LowLim column after writing the column header,
                                ' once the first (number) limit is encountered
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("F2:F2002").Select
                                    Range("F2:F2002").NumberFormat = "0.000000000"
                                End If
                                 
                           End If
                       Next copycount
                    End If
                   
                   'If c_ltl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                   'Copy Cold Temp, Final Test Upper Limit to the "Test Limits" worksheet
                   If c_utl_col > 0 Then
                      For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Cold Limits").Cells(copycount, c_utl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "G").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "G").Value = current_text
                               
                                If (IsNumeric(current_text)) Then
                                    If (copycount = 2) Then
                                        paste_workbook.Worksheets("Test_Limits").Activate
                                        ' Set formatting for first 2000 limits
                                        Range("G2:G2002").Select
                                        Range("G2:G2002").NumberFormat = "0.000000000"
                                    End If
                                End If
                                
                           End If
                      Next copycount
                      
                   End If
                   
                   'If c_utl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                 
                   'Copy Hot Temp, Final Test Lower Limit to the "Test Limits" worksheet
                    If h_ltl_col > 0 Then
                       For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Hot Limits").Cells(copycount, h_ltl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "H").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "H").Value = current_text
                                ' In certain instance, the formating of the limits from product engineering may contain
                                ' too many digits past the decimal point. So manage formating here.
                                ' Format the numbers in the RT_LowLim column after writing the column header,
                                ' once the first (number) limit is encountered
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("H2:H2002").Select
                                    Range("H2:H2002").NumberFormat = "0.000000000"
                                End If
                                 
                           End If
                       Next copycount
                    End If
                   
                   'If h_ltl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                   'Copy Hot Temp, Final Test Upper Limit to the "Test Limits" worksheet
                   If h_utl_col > 0 Then
                      For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Hot Limits").Cells(copycount, h_utl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "I").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "I").Value = current_text
                               
                                If (IsNumeric(current_text)) Then
                                    If (copycount = 2) Then
                                        paste_workbook.Worksheets("Test_Limits").Activate
                                        ' Set formatting for first 2000 limits
                                        Range("I2:I2002").Select
                                        Range("I2:I2002").NumberFormat = "0.000000000"
                                    End If
                                End If
                                
                           End If
                      Next copycount
                      
                   End If
                   
                   'If h_utl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                    'Copy Room Temp, QC Test Lower Limit to the "Test Limits" worksheet
                    If qcr_ltl_col > 0 Then
                       For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Production Limits").Cells(copycount, qcr_ltl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "J").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "J").Value = current_text
                               
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("J2:J2002").Select
                                    Range("J2:J2002").NumberFormat = "0.000000000"
                                End If
                           End If
                       Next copycount
                       
                    End If
                
                    'If qcr_ltl_col > 0 Then
                    '        mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                    '         columnshift = columnshift + 1
                    'End If

                    'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                    'Copy Room Temp, QC Test Upper Limit to the "Test Limits" worksheet
                    If qcr_utl_col > 0 Then
                       For copycount = 1 To count + 1
                           current_text = src_data.Worksheets("Production Limits").Cells(copycount, qcr_utl_col).Value
                           
                           If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "K").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                           
                           'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                 paste_workbook.Worksheets("Test_Limits").Cells(copycount, "K").Value = current_text
                                   
                                 If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("K2:K2002").Select
                                    Range("K2:K2002").NumberFormat = "0.000000000"
                                End If
                           End If
                       Next copycount
                       
                    End If
                
                    'If qcr_utl_col > 0 Then
                    '        mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                    '        columnshift = columnshift + 1
                    'End If

                    'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                    'Copy Cold Temp, QC Test Lower Limit to the "Test Limits" worksheet
                    If qcc_ltl_col > 0 Then
                      For copycount = 1 To count + 1
                          current_text = src_data.Worksheets("Cold Limits").Cells(copycount, qcc_ltl_col).Value
                          
                          If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "L").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                          
                          'If IsNumeric(current_text) Then
                            formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                            current_text = CSng(formated_string)
                          
                          
                            paste_workbook.Worksheets("Test_Limits").Cells(copycount, "L").Value = current_text
                               
                            If (copycount = 2) Then
                                paste_workbook.Worksheets("Test_Limits").Activate
                                ' Set formatting for first 2000 limits
                                Range("L2:L2002").Select
                                Range("L2:L2002").NumberFormat = "0.000000000"
                            End If
        
                        End If
                      Next copycount
                   End If
                   
                   'If qcc_ltl_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   '     columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                    
                    'Copy Cold Temp, QC Test Upper Limit to the "Test Limits" worksheet
                    If qcc_utl_col > 0 Then
                      For copycount = 1 To count + 1
                          current_text = src_data.Worksheets("Cold Limits").Cells(copycount, qcc_utl_col).Value
                          
                          If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "M").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                          
                          'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "M").Value = current_text
                               
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("M2:M2002").Select
                                    Range("M2:M2002").NumberFormat = "0.000000000"
                                End If
                           End If
                      Next copycount
                   End If
                   
                   'If qcc_utl_col > 0 Then
                    '    mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                    '    columnshift = columnshift + 1
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                   'Copy Hot Temp, QC Test Lower Limit to the "Test Limits" worksheet
                   If qch_ltl_col > 0 Then
                      For copycount = 1 To count + 1
                          current_text = src_data.Worksheets("Hot Limits").Cells(copycount, qch_ltl_col).Value
                          
                          If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "O").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                          'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "O").Value = current_text
                               
                                If (copycount = 2) Then
                                    paste_workbook.Worksheets("Test_Limits").Activate
                                    ' Set formatting for first 2000 limits
                                    Range("O2:O2002").Select
                                    Range("O2:O2002").NumberFormat = "0.000000000"
                                End If
                           End If
                      Next copycount
                      
                   End If
                   
                   'If  qch_ltl_col > 0 Then
                        'mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                        
                   'End If
                
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
   
                   'Copy Hot Temp, QC Test Upper Limit to the "Test Limits" worksheet
                   If qch_utl_col > 0 Then
                      For copycount = 1 To count + 1
                          current_text = src_data.Worksheets("Hot Limits").Cells(copycount, qch_utl_col).Value
                          
                          If copycount = 1 Then                 ' if copycount = 1 copy column header to Test Limits worksheet
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "P").Value = current_text
                           
                           ElseIf IsNumeric(current_text) Then   ' else make sure the test limit is a number before copying to the Test Limits worksheet
                          
                                ' Format number will convert a number to a string but provides the ability to dictate digits to the right of the decimal point.
                                ' The CSNG converts the string to a single with a range of acceptable values - > -3.4028E38  to 3.4028E38
                          'If IsNumeric(current_text) Then
                                formated_string = FormatNumber(current_text, 9, vbFalse, vbFalse, vbFalse)
                                current_text = CSng(formated_string)
                                
                                paste_workbook.Worksheets("Test_Limits").Cells(copycount, "P").Value = current_text
                                If (copycount = 2) Then
                                        paste_workbook.Worksheets("Test_Limits").Activate
                                        ' Set formatting for first 2000 limits
                                        Range("P2:P2002").Select
                                        Range("P2:P2002").NumberFormat = "0.000000000"
                                End If
                           End If
                      Next copycount
                    End If
                
                   'If qch_utl_col > 0 Then
                        'mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                        
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP

                        'COPY UNITS'
                    If units_col > 0 Then
                     For copycount = 1 To count + 1
                         current_text = src_data.Worksheets(prod_names_alt).Cells(copycount, units_col).Value
                          
                          If (IsEmpty(current_text) = False) Then
                               paste_workbook.Worksheets("Test_Limits").Cells(copycount, "Q").Value = current_text
                          End If
                   
                    Next copycount
                    
                    End If
                
                   'If units_col > 0 Then
                   '     mytest = copy_limits_to_utilities(testname_col, columnshift, NUM_PROD_LIMITS)
                   'End If
                   
                   'BSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRBSRCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCPCCP
                
                ' Do not reference the entire path to close the workbook
                paste_workbook.Worksheets("Test_Limits").Columns("A:Q").AutoFit
                Workbooks(direct_filename).Close
                paste_workbook.Worksheets("PDS Utilities").Activate
            End If
            
            ' Now that data has been transferred to the Test_Limits worksheet, identify duplicate test names
            ' if they exist and inform the user.
            ' ALL test names must be unique for proper transfer of limits to pds files
            ' Test_Names column is 2nd (B) column in Test_Limits work sheet.
    
            Dim duplicate_tests As Integer
    
            ' Search the "Test_Limits" worksheet starting from row+2 to row=count at column 2 for duplicate test names
            duplicate_tests = Check_4Duplicate_Test_Names("Test_Limits", 2, count, 2)
                            
           End If
                    
    End If
    
    
End Sub
     
Public Function worksheet_exists(sheetname As String) As Boolean
 
 Dim TempSheetName As String
 Dim sheet As Worksheet
  
 TempSheetName = UCase(sheetname)
 
 worksheet_exists = False
 
 For Each sheet In Worksheets
   If TempSheetName = UCase(sheet.Name) Then
     worksheet_exists = True
     Exit Function
  End If
 

 Next sheet

End Function

Function copy_limits_to_utilities(ByVal src_column_index As Integer, ByVal dest_column_index As Integer, ByVal num_limits As Integer) As Boolean

Dim copycount As Integer
Dim current_text As String

  '  If src_column_index > 0 Then
  '      For copycount = 1 To num_limits + 1
  '          current_text = src_data.Worksheets("Production Limits").Cells(copycount, src_column_index).Value
  '         paste_workbook.Worksheets("Test_Limits").Cells(copycount, dest_column_index).Value = current_text
  '     Next copycount
  '
  '  End If
    
    
   'copy_limits_to_utilities = True
End Function

Public Function Is_File_Open(filename As String)

'Dim filenum As Integer, errnum As Integer

    'On Error Resume Next   ' Turn error checking off.
    'filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    'Open filename For Input Lock Read As #filenum
    'Close filenum          ' Close the file.
    'errnum = Err           ' Save the error number that occurred.
    'On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    'Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        'Case 0
         'IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        'Case 70
            'IsFileOpen = True

        ' Another error occurred.
        'Case Else
            'Error errnum
    'End Select


End Function



' Search the evaluated worksheet starting from start_row to stop_row at eval_column for duplicate test names
Public Function Check_4Duplicate_Test_Names(eval_worksheet As String, start_row As Integer, stop_row As Integer, eval_column As Integer)

 Dim compare_stg As String, search_stg As String
 Dim search_index As Integer, compare_index As Integer, results As Integer
 Dim dupeList As Object
 Set dupeList = CreateObject("System.Collections.ArrayList")
 Dim itemInDupe As Integer
 
    Check_4Duplicate_Test_Names = 0
    
    For search_index = start_row To stop_row
    
         search_stg = ThisWorkbook.Worksheets(eval_worksheet).Cells(search_index, eval_column).Value
         
         For compare_index = (search_index + 1) To stop_row
         
            compare_stg = ThisWorkbook.Worksheets(eval_worksheet).Cells(compare_index, eval_column).Value
             
             results = 2  ' default to a value not returned by the function that is an integer
             results = StrComp(search_stg, compare_stg, vbTextCompare)
             If (results = 0) And Not search_stg = "" Then
                
                Check_4Duplicate_Test_Names = Check_4Duplicate_Test_Names + 1
                dupeList.Add search_stg
               
             End If
             
        Next compare_index
    Next search_index
    
    If Check_4Duplicate_Test_Names > 0 Then
        MsgBox "There were ''" & Check_4Duplicate_Test_Names & "'' duplicate test name(s) found!", vbCritical, "Duplicate Test Names"
        
        For itemInDupe = 0 To dupeList.count - 1
            MsgBox "Duplicate Test: ''" & dupeList.Item(itemInDupe) & "''", vbCritical, "Dupe Test Name"
        Next itemInDupe
        
    End If
End Function
