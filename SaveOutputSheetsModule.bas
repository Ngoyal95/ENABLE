Attribute VB_Name = "SaveOutputSheetsModule"
Sub SaveOutputSheets()
'Sub to save 'Combined' and 'Outputs' and then delete them
Dim SaveName, flpt, strDate As String
    
strDate = Format(Date, "mm-dd-yyyy")

    If WorksheetExists("Combined") = False Or WorksheetExists("Output") = False Then
        MsgBox "Outputs and/or Combined sheets not generated! Run the program."
    Else
        'SaveName = InputBox("Enter name to save the output sheets file as")
        SaveName = "Compiled Data_" & strDate
        flpt = OutputSheetsLoc & "\" & SaveName & ".xlsx"
        Sheets(Array("Combined", "Output")).Move
        ActiveWorkbook.SaveAs FileName:=flpt, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Application.ActiveWorkbook.Close False
    End If
End Sub
