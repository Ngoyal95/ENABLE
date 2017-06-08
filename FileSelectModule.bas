Attribute VB_Name = "FileSelectModule"
Sub FileSelectAndImport()

Dim vaFiles As Variant
Dim i As Long
Dim wbkToCopy As Workbook

vaFiles = Application.GetOpenFilename(MultiSelect:=True)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

If IsArray(vaFiles) Then
    For i = LBound(vaFiles) To UBound(vaFiles)
        Set wbkToCopy = Workbooks.Open(FileName:=vaFiles(i))
        '
        wbkToCopy.Sheets.Item(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).Name = wbkToCopy.Name
        '
        wbkToCopy.Close savechanges:=False
    Next i
End If

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

