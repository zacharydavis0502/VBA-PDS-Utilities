Attribute VB_Name = "Reset_Project"

' The Reset Project button is intended to remove all worksheets
' from a project in order to re-read Product Engineering limits,
' re-import pds files, re-transfer limits to pds files and re-export
' pds files to pds format.

Sub Reset_Project()
  
  Dim response As VbMsgBoxResult
  Dim num_sheets As Integer
  Dim do_nothing As Boolean
  Dim sheet_index As Integer
  Dim sheet_name As String
  Dim util_sheet, read_me As String
  Dim sheet_count As Integer
  Dim delete_sheet As Boolean
  
  
  Set my_workbook = ActiveWorkbook
  util_sheet = "PDS Utilities"
  read_me = "Read_Me"
  
  
  response = MsgBox("Reseting the project will remove ALL worksheets except PDS Utilities and Read Me sheets. Are you sure!!", vbYesNo, "RESET PROJECT?")
  
  Select Case response
            Case vbYes
                num_sheets = ThisWorkbook.Worksheets.count
               
                Dim compare As Integer
                delete_sheet = False
                ' If more than 2 worksheets then there are sheets in addition
                ' to "PDS Utilities" and "Read Me" worksheets. Remove these worksheets.
                If num_sheets > 2 Then
                     For sheet_index = 1 To num_sheets
                       sheet_name = my_workbook.Sheets(sheet_index).Name
                       If (UCase(sheet_name) = UCase(util_sheet)) Then
                            delete_sheet = False
                       ElseIf (UCase(sheet_name) = UCase(read_me)) Then
                            delete_sheet = False
                       Else
                            delete_sheet = True
                            Sheets(sheet_index).Delete
                            num_sheets = num_sheets - 1
                            sheet_index = sheet_index - 1
                       End If
                       If sheet_index = num_sheets Then Exit For
                     Next sheet_index
                End If
                MsgBox ("Project Reset Completed")
             Case Else
                MsgBox ("Project Reset Cancelled")
  End Select

End Sub
