# CayoAccessUpload- File is in Visual Basic

Option Compare Database
Option Explicit
Public Sub ReadParsedData(ByVal FilePath As String) ', ByRef txtOutput As TextBox)[I think this is reading and assigning the text that I input on the Form as FilePath]
   Function Impo_allExcel() '[The loop function I want to run]

   'Dim myExcel As New Excel.Application
   Dim db As Database
   Dim I As Integer
   Dim SheetName As String
   Dim StepName As String
   Dim FileName As String
   'Dim mySheet As Excel.Worksheet
   
   ChDir (FilePath)
   FileName = Dir() '[Defining FileName as the last part of the FilePath]
   
   'On Error Resume Next
   
   Do While FileName <> ""
    If myfile Like "*.xls*" Then '[Loop while there is a File that is an xls]
   
    'Clean out all the records from the temporary tables
    'Deleting the records from "tmp" tables will delete all records in all of child "tmp" tables
    'because of Cascade Delete Related Records setting
   
    On Error Resume Next
    CurrentDb.Execute ("Delete * from tmpFocal")

    'Transfer records from the data parser spreadsheets
    On Error GoTo TransferError
    'Focal Data
    StepName = "Transfering Records to tmpFocal"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocal", FilePath & FileName, True, "Focal!A:F"
    StepName = "Transfering Reccords to tmpFocalBehavior"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalBehavior", FilePath & FileName, True, "FocalBehavior!A:O"
    StepName = "Transfering Reccords to tmpFocalAdLib"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalAdlib", FilePath & FileName, True, "FocalAdLib!A:M"
    StepName = "Transfering Reccords to tmpFocalConsort"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalConsort", FilePath & FileName, True, "FocalConsort!A:K"
    StepName = "Transfering Reccords to tmpFocalPause"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpFocalPause", FilePath & FileName, True, "FocalPause!A:I"
    StepName = "Transfering Records to tmpPointScan"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpPointScan", FilePath & FileName, True, "PointScan!A:H"
    StepName = "Transfering Records to tmpPointScanProximity"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tmpPointScanProximity", FilePath & FileName, True, "PointScanProximity!A:I"
   
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
    Exit Function
TransferError:
    MsgBox ("Error " & StepName & vbCrLf & Err.Description & vbCrLf & "Transfer Aborted")
    'CurrentDb.Execute ("Delete * from tmpFocals")
    Exit Function
AppendError:
    MsgBox ("Error " & StepName & vbCrLf & Err.Description & vbCrLf & "Appending of Records Aborted")
    Exit Function
    End Function
    
End If
FileName = Dir() '[End the loop]
Loop
End Function
