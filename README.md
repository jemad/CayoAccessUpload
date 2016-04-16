Option Compare Database
Option Explicit
Sub LoopAllExcelFilesInFolder() 'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them

Dim myPath As String
Dim myFile As String
Dim myExtension As String


'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      FileName = (myPath & myFile)
    
  Public Sub ReadParsedData(ByVal FolderName As String) ', ByRef txtOutput As TextBox)

   'Dim myExcel As New Excel.Application
   Dim db As Database
   Dim I As Integer
   Dim SheetName As String
   Dim StepName As String
   'Dim mySheet As Excel.Worksheet
   
   
   'On Error Resume Next
   
   'Clean out all the records from the temporary tables
   'Deleting the records from "tmp" tables will delete all records in all of child "tmp" tables
   'because of Cascade Delete Related Records setting
   
   On Error Resume Next
   CurrentDb.Execute ("Delete * from tmpFocal")

   'Transfer records from the data parser spreadsheets
   On Error GoTo TransferError
   'Focal Data
   StepName = "Transfering Records to tmpFocal"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocal", FileName, True, "Focal!A:F"
   StepName = "Transfering Reccords to tmpFocalBehavior"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalBehavior", FileName, True, "FocalBehavior!A:O"
   StepName = "Transfering Reccords to tmpFocalAdLib"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalAdlib", FileName, True, "FocalAdLib!A:M"
   StepName = "Transfering Reccords to tmpFocalConsort"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalConsort", FileName, True, "FocalConsort!A:K"
   StepName = "Transfering Reccords to tmpFocalPause"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalPause", FileName, True, "FocalPause!A:I"
   StepName = "Transfering Records to tmpPointScan"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpPointScan", FileName, True, "PointScan!A:H"
   StepName = "Transfering Records to tmpPointScanProximity"
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpPointScanProximity", FileName, True, "PointScanProximity!A:I"
   
'N.B. Transfer errors can occur if there are blank line after the data in the Excel spreadsheets
   
   'Append the records from the temporary tables into the permanent tables
   On Error GoTo AppendError
   
   StepName = "Deleting tmpFocal Records from tblFocal"
   CurrentDb.Execute "Delete-tmpFocal-from-tblFocal", dbFailOnError
   
   StepName = "Appending Records to tblFocal"
   CurrentDb.Execute "Import-Focal-Step1-Append-Focal", dbFailOnError
   
   StepName = "Appending Records to tblFocalBehavior"
   CurrentDb.Execute "Import-Focal-Step2-Append-FocalBehavior", dbFailOnError
   
   StepName = "Appending Records to tblFocalAdLib"
   CurrentDb.Execute "Import-Focal-Step3-Append-FocalAdLib", dbFailOnError
   
   StepName = "Appending Records to tblFocalConsort"
   CurrentDb.Execute "Import-Focal-Step4-Append-FocalConsort", dbFailOnError
   
   StepName = "Appending Records to tblFocalPause"
   CurrentDb.Execute "Import-Focal-Step5-Append-FocalPause", dbFailOnError
   
   StepName = "Appending Records to tblPointScan"
   CurrentDb.Execute "Import-Focal-Step6-Append-PointScan", dbFailOnError
   
   StepName = "Appending Records to tblPointScanProximity"
   CurrentDb.Execute "Import-Focal-Step7-Append-PointScanProximity", dbFailOnError

Exit_ReadParsedData:
   MsgBox ("Data Was Imported Successfully From " & FileName)
   Exit Sub
TransferError:
   MsgBox ("Error " & StepName & vbCrLf & Err.Description & vbCrLf & "Transfer Aborted")
   'CurrentDb.Execute ("Delete * from tmpFocals")
   Exit Sub
AppendError:
   MsgBox ("Error " & StepName & vbCrLf & Err.Description & vbCrLf & "Appending of Records Aborted")
   Exit Sub
End Sub
  

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
