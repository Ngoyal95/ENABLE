Attribute VB_Name = "ExtractNumberFromText"
Function Extract_Number_from_Text(Phrase As String) As Integer
Dim Length_of_String As Integer
Dim Current_Pos As Integer
Dim Temp As String

Length_of_String = Len(Phrase)
Temp = ""

For Current_Pos = 1 To Length_of_String

If (Mid(Phrase, Current_Pos, 1) = "-") Then
  Temp = Temp & Mid(Phrase, Current_Pos, 1)
End If


If (Mid(Phrase, Current_Pos, 1) = ".") Then
 Temp = Temp & Mid(Phrase, Current_Pos, 1)
End If


If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
    Temp = Temp & Mid(Phrase, Current_Pos, 1)
    
 End If

Next Current_Pos

If Len(Temp) = 0 Then
    Extract_Number_from_Text = 0
Else
    Extract_Number_from_Text = CInt(Temp)
End If

End Function

Sub test()
Dim testval As Double

testval = Extract_Number_from_Text("Left Renal Mass; TL")
MsgBox testval

End Sub
