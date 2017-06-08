Attribute VB_Name = "RECISTWordExportationModule"
Sub RECISTWordExport(ws As Worksheet)
'https://anictteacher.files.wordpress.com/2011/11/vba-error-462-explained-and-resolved.pdf
'Above link shows how to make it not need a reset after performing a word export

Application.ScreenUpdating = False 'Speeds up operation

'initializing and opening word
'Dim objApplication As Object
Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim RECISTDocPath, strPatientName, strDate, strTime As String
Dim SaveDocPath As String
Dim bWeStartedWord As Boolean

'Populating the table
Dim i, InsertVar, PopVal, NumberingFlag As Integer
Dim LesionNumber As Integer
Dim CellText As String

'Initialize variables
PopVal = 0


'=================================
'   launch word Application
'=================================
'Application.DisplayAlerts = True

On Error Resume Next
Set wdApp = GetObject(, "Word.Application") 'Try to get open word instance so does not create new instance
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application") 'new instance of word
End If
On Error GoTo 0


SaveDocPath = WordDocLoc 'global var, set in the FolderSetup module
strPatientName = ws.Cells(3, PatNameLoc)
RECISTDoc = Application.ActiveWorkbook.path & "\RECISTForm.docx"
strDate = Format(Date, "mm-dd-yyyy")
strTime = Format(Time, "hh.nn AM/PM")

On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
On Error GoTo 0
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
    bWeStartedWord = True
End If

wdApp.Visible = False  'optional, not required
Set wdDoc = wdApp.Documents.Open(RECISTDoc)

With wdDoc
    '================================
    'Populate the RECIST sheet
    '================================
    For i = 3 To EOCE - 1
    LesionNumber = Extract_Number_from_Text(ws.Cells(i, DescripLoc).value)
    If LesionNumber <> 0 Then
        NumberingFlag = 0
    ElseIf LesionNumber = 0 Then
        NumberingFlag = 1
    End If
    Next i
    
    If NumberingFlag = 0 And NLTrack = 0 Then 'Dont order with New lesions
        'Use Numbering
        InsertVar = 2
        For i = 3 To EOCE - 1
            CellText = ws.Cells(i, DescripLoc).value
            LesionNumber = Extract_Number_from_Text(CellText)
            CellText = ws.Cells(i, TargetLoc)

            If LesionNumber > 0 Then
                'this next if block is used to determine what row of the RECIST form table to print to
                If (InStr(1, CellText, "Target") > 0) And (InStr(1, CellText, "Non-") = 0) Then 'Target measurement
                    PopVal = InserVar + LesionNumber + 1
                ElseIf (InStr(1, CellText, "Non-Target") > 0) Then 'Non-Target measurement
                    PopVal = InsertVar + TTrack + LesionNumber - 1
                End If
            ElseIf LesionNumber = 0 Then
                PopVal = PopVal + 1
            End If

            If (TypeName(ws.Cells(i, RECISTDiaLoc).value) = "Double") Then
                If LesionNumber <> 0 Then
                    .Tables(3).cell(PopVal, 1).Range.Text = LesionNumber 'populate Lesion#
                End If
                .Tables(3).cell(PopVal, 5).Range.Text = Modality(0) 'Populate measurement modality of most recent CT
                .Tables(3).cell(PopVal, 2).Range.Text = ws.Cells(i, FllwUpLoc) & " / " & ws.Cells(i, DescripLoc) 'populate Descriptions

                CellText = ws.Cells(i, TargetLoc) 'store the text to check if Target or Non-Target
                If (InStr(1, CellText, "Target") > 0) And (InStr(1, CellText, "Non") = 0) Then
                    'Target measurement
                    .Tables(3).cell(PopVal, 3).Range.Text = "T" 'Populate Target field
                ElseIf (InStr(1, CellText, "Non-Target") > 0) Then
                    'Non-Target measurement
                    .Tables(3).cell(PopVal, 3).Range.Text = "NT" 'Populate target field
                End If

                .Tables(3).cell(PopVal, 7).Range.Text = ws.Cells(i, SeriesLoc) & "/" & ws.Cells(i, SliceLoc) 'Populate Series/Image#
                .Tables(3).cell(PopVal, 8).Range.Text = Round(ws.Cells(i, RECISTDiaLoc).value, 1) 'populate RECIST Diameters

                CellText = ws.Cells(i, StdDescpLoc + 1)
                .Tables(3).cell(PopVal, 9).Range.Text = Left(CellText, InStr(CellText, " ")) 'populate Date
            End If
        Next i
    Else
        'Do not use numbering, insert as they appear in bookmark table
        InsertVar = 2
        For i = 3 To EOCE - 1
            If (TypeName(ws.Cells(i, RECISTDiaLoc).value) = "Double") Then

                'wdDoc.Tables(3).cell(InsertVar, 1).Range.Text = InsertVar - 1 'populate Lesion#
                .Tables(3).cell(InsertVar, 5).Range.Text = Modality(0) 'Populate measurement modality of most recent CT
                .Tables(3).cell(InsertVar, 2).Range.Text = ws.Cells(i, FllwUpLoc) & " / " & ws.Cells(i, DescripLoc) 'populate Descriptions

                CellText = ws.Cells(i, TargetLoc) 'store the text to check if Target or Non-Target
                If (InStr(1, CellText, "Target") > 0) And (InStr(1, CellText, "Non") = 0) Then
                    'Target measurement
                    .Tables(3).cell(InsertVar, 3).Range.Text = "T" 'Populate Target field
                ElseIf (InStr(1, CellText, "Non-Target") > 0) Then
                    'Non-Target measurement
                    .Tables(3).cell(InsertVar, 3).Range.Text = "NT" 'Populate target field
                End If

                .Tables(3).cell(InsertVar, 7).Range.Text = ws.Cells(i, SeriesLoc) & "/" & ws.Cells(i, SliceLoc) 'Populate Series/Image#
                .Tables(3).cell(InsertVar, 8).Range.Text = Round(ws.Cells(i, RECISTDiaLoc).value, 1) 'populate RECIST Diameters

                CellText = ws.Cells(i, StdDescpLoc + 1)
                .Tables(3).cell(InsertVar, 9).Range.Text = Left(CellText, InStr(CellText, " ")) 'populate Date

                InsertVar = InsertVar + 1
            ElseIf (TypeName(Cells(i, TargetLoc).value) <> "Double") Or (IsEmpty(ws.Cells(i, RECISTDiaLoc)) = True) Then
                Exit For
            End If
        Next i
    End If

    If InstanceCounter > 1 Then
        '5Values
        .Tables(4).cell(1, 2).Range.Text = ws.Cells(2, LastCol + 1)
        .Tables(4).cell(2, 2).Range.Text = TLBS
        .Tables(4).cell(3, 2).Range.Text = BestResponse
        .Tables(4).cell(4, 2).Range.Text = Round(100 * ((CTLS - BestResponse) / BestResponse), 0)
        .Tables(4).cell(5, 2).Range.Text = ws.Cells(2, LastCol + 2)
        
        If UserInterface.ResponseDet = True Then
            'Response
            .Tables(4).cell(6, 2).Range.Text = ResponseT
            .Tables(4).cell(7, 2).Range.Text = ResponseNT
            .Tables(4).cell(8, 2).Range.Text = OverallResponse
        End If
    Else
        '5Values
        .Tables(4).cell(1, 2).Range.Text = ws.Cells(2, LastCol + 1)
        .Tables(4).cell(2, 2).Range.Text = ws.Cells(2, LastCol + 1)
        .Tables(4).cell(3, 2).Range.Text = " - "
        .Tables(4).cell(4, 2).Range.Text = " - "
        .Tables(4).cell(5, 2).Range.Text = " - "
        
        'Response
        .Tables(4).cell(6, 2).Range.Text = " - "
        .Tables(4).cell(7, 2).Range.Text = " - "
        .Tables(4).cell(8, 2).Range.Text = " - "
    End If

    '==================================================================================
    'Generic info: "Measured By", followup or not, date, patient name, protocol #, MRN
    '==================================================================================
    'Follow up or not:
    If InstanceCounter > 1 Then
        .Tables(2).cell(1, 9).Range.Text = "(X)"
    Else
        .Tables(2).cell(1, 6).Range.Text = "(X)"
    End If
    
    'Protocol #
    .Tables(2).cell(1, 2).Range.Text = ProtocolNumber
    
    'Course # (NOT IN TABLE)
    '.Tables(2).cell(1,4).Range.Text = ""
    
    'Measured by
    .Tables(5).cell(2, 2).Range.Text = ConcatNames(ws.Range(ws.Cells(3, CreatorLoc), ws.Cells(EOCE, CreatorLoc))) 'who it was measured by
    
    'Date measured
    'Assume that it is same day as the most recent CT (ideal situation)
    .Tables(5).cell(1, 2).Range.Text = Left(ws.Cells(3, StdDescpLoc + 1), InStr(ws.Cells(3, StdDescpLoc + 1), " ")) 'populate Date
    
    'Patient identifier & MRN
    .Tables(6).cell(2, 1).Range.Text = ws.Cells(3, PatNameLoc) & vbNewLine & MedicalRecordNumber
End With

'=================================
'   Save and exit word Application
'=================================
wdDoc.SaveAs2 SaveDocPath & strPatientName & "_" & strDate & "_" & strTime & ".docx"  'Saves as a new doc
'wdDoc.SaveAs2 SaveDocPath & ActiveSheet.Name & "_" & strDate & ".docx" 'Saves as a new doc (for PHI in images)
wdDoc.Close
ExitPoint:
    If bWeStartedWord Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    
Application.ScreenUpdating = True

End Sub

