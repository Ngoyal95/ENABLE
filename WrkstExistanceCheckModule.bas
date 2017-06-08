Attribute VB_Name = "WrkstExistanceCheckModule"
Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function
