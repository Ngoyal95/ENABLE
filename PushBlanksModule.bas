Attribute VB_Name = "PushBlanksModule"
Sub PushBlanks(ws As Worksheet)
'Push down blank rows

With ws.UsedRange
    rowcount = .Rows.count
    colcount = .Columns.count
    For i = rowcount To 1 Step -1
        k = 0
        For j = 1 To colcount
            If .Value2(i, j) <> "" Then
                k = 1
                Exit For
            End If
    
        Next j
        If k = 0 Then
            .Rows(i).Delete Shift:=xlUp
        End If
    Next i
End With
End Sub
