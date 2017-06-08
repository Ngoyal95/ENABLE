Attribute VB_Name = "LMFormatter"
Sub LMWkbkFormatter(wkbk As Workbook)
'This runs on the entire workbook

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim ws As Worksheet

On Error Resume Next
Set ws = wkbk.Worksheets.Item(1)
On Error GoTo 0
'make sure we have at least one visible sheet


If Not ws Is Nothing Then
    For Each ws In wkbk.Worksheets
        If Not ws.Name = "Main" And Not ws.Name = "Output" And Not ws.Name = "Combined" Then
            With ws
                LastCol = .Cells(1, ws.Columns.count).End(xlToLeft).Column
                For StartNumber = 0 To InstanceCounter - 1
                'Goto head header loc, sec equal to the HeaderInstanceLocation(StartNumber)+1 B cell value
                .Cells(StudyInstanceLocations(StartNumber), StdDescpLoc) = .Cells((StudyInstanceLocations(StartNumber) + 1), (StdDescpLoc + 1)).value
                Next StartNumber
                '==========================
                'Add a blank row after row 1 ONLY UNTIL LASTCOL+9 where the patient identifier data starts
                    .Range(.Cells(2, 1), .Cells(LastRow, LastCol - 3)).Cut .Range("A3") 'Note,  use -3 because the lastCol will include the 3 patient identifier columns at the far right of spreadsheet
                End With

            '============================================================================
            'Save the sheet in new file as .xls, delete the sheet from the Interface wkbk
            '============================================================================
            ws.Copy
            SaveName = Left(ws.Name, Application.WorksheetFunction.Find(".", ws.Name) - 1)
            flpt = LabmatrixLoc & SaveName
            ActiveWorkbook.SaveAs FileName:=flpt, FileFormat:=xlExcel8, CreateBackup:=False
            Application.ActiveWorkbook.Close False
        End If
    Next ws
End If

End Sub

