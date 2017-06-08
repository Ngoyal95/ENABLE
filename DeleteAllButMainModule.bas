Attribute VB_Name = "DeleteAllButMainModule"
Sub DeleteAllButMain()
    ' Ref http://www.ozgrid.com/forum/showthread.php?t=17680
Application.ScreenUpdating = False 'Speeds up operation
Application.DisplayAlerts = False

Dim ws As Worksheet

On Error Resume Next
Set ws = Worksheets.Item(2)
On Error GoTo 0
'make sure we have at least one visible sheet
If Not ws Is Nothing Then

    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name = "Main" Then ws.Delete  'Does not delete "Main"
    Next ws
End If
'delete all the others

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


