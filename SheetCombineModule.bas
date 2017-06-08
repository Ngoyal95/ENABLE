Attribute VB_Name = "SheetCombineModule"
Sub CombineSheets()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim CopyRng As Range
    Dim Flag As Integer
    Dim RunSum As Long
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Delete the summary sheet if it exists.
    Application.DisplayAlerts = False
    On Error Resume Next
    Workbooks(InterfaceVersion).Worksheets("Combined").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add a new summary worksheet.
    If ThisWorkbook.Sheets.count > 1 Then
        'ActiveWorkbook.Worksheets.Add After:=Worksheets("Main")
        'ActiveWorkbook.Sheets.Item(2).Name = "Combined"
        Sheets.Add(After:=Sheets("Main")).Name = "Combined"
        Set DestSh = Sheets("Combined")
    
        ' Loop through all worksheets and copy the data to the summary worksheet.
        Flag = 0
        RunSum = 0
        For Each sh In Workbooks(InterfaceVersion).Worksheets
            If sh.Name <> DestSh.Name And sh.Name <> "Main" And sh.Name <> "Output" And InStr(1, sh.Name, "LM Copy") = 0 Then
                sh.Activate
                
                If Flag = 0 Then
                    sh.UsedRange.Copy Sheets("Combined").Cells(1, "A")
                    RunSum = sh.Range("B65000").End(xlUp).row + 5
                    Flag = 1
                ElseIf Flag = 1 Then
                    sh.UsedRange.Copy Sheets("Combined").Cells(RunSum, "A")
                    RunSum = RunSum + sh.Range("B65000").End(xlUp).row + 4
                End If
            End If
        Next

ExitTheSub:
    
        Application.GoTo DestSh.Cells(1)
    
        ' AutoFit the column width in the summary sheet.
        DestSh.Columns.AutoFit
End If
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

