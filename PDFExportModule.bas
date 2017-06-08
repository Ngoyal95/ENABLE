Attribute VB_Name = "PDFExportModule"
 'Ref: https://support.microsoft.com/en-us/kb/177760
 'Sub Procedure to Run an Existing Microsoft Word Macro
Sub PDF()

Dim objWord As Object
'We need to continue through errors since if Word isn't
'open the GetObject line will give an error
On Error Resume Next
Set objWord = GetObject(, "Word.Application")

'We've tried to get Word but if it's nothing then it isn't open
If objWord Is Nothing Then
Set objWord = CreateObject("Word.Application")
End If
objWord.Visible = False

'It's good practice to reset error warnings
On Error GoTo 0

'Open your document and ensure its visible and activate after openning
'objWord.Documents.Open Worksheets.Item(1).Cells(3, "A").value & "PDFC.docm"

objWord.Documents.Open Application.ActiveWorkbook.path & "\PDFC.docm"
objWord.Run "Project.Module1.ChangeDocsToTxtOrRTFOrHTML"
'objWord.Activate

Set objWord = Nothing


End Sub



    




