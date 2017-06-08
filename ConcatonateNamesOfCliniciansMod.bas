Attribute VB_Name = "ConcatonateNamesOfCliniciansMod"
Function ConcatNames(RowRange As Range) As String
  Dim x As Long, CellVal As String, ReturnVal As String, Result As String
  Const Delimiter = ", "
  For x = 1 To RowRange.count
    ReturnVal = RowRange(x).value
    If Len(RowRange(x).value) Then If InStr(Result & Delimiter, Delimiter & ReturnVal & Delimiter) = 0 Then Result = Result & Delimiter & ReturnVal
  Next
  ConcatNames = Mid(Result, Len(Delimiter) + 1)
End Function
