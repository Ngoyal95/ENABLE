Attribute VB_Name = "DeleteSheetsModule"
Sub DeleteSheets()
    ' Ref http://www.ozgrid.com/forum/showthread.php?t=17680
Application.ScreenUpdating = False 'Speeds up operation
Dim ws As Worksheet

On Error Resume Next
Set ws = Worksheets.Item(2)
On Error GoTo 0
'make sure we have at least one visible sheet
If Not ws Is Nothing Then
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name = "Main" And Not ws.Name = "Output" And Not ws.Name = "Combined" Then ws.Delete  'Does not delete "Main"
    Next ws
    Application.DisplayAlerts = True
End If
'delete all the others

Application.ScreenUpdating = True
End Sub
