Attribute VB_Name = "RECISTCalcModule"
Sub RECISTCalc(ws As Worksheet)
'Takes args: the current worksheet

'===========
'Variables:
'===========
Dim StartNumber As Integer 'Startnumber is used in the forloop, LastRow stores the index of the last row filled with data (Col B is used for Bookmark Table)
Dim CellText As String 'used to store the cell's string
Dim ArrayPos As Integer 'used for entering data into arrays
Dim iVal As Double 'Used when finding the BestResponse
Dim rng1, rng2 As Range
'Dim SutdyInstanceLocations() As Integer

'==========
'CODE:
'==========
InstanceCounter = Application.WorksheetFunction.CountIf(ws.Range(ws.Cells(1, StdDescpLoc), ws.Cells(LastRow, StdDescpLoc)), "STUDY INSTANCE UID*") 'Count the number of studies using Col A where "STUDY INSTANCE" appears

'==================
'Initialize values:
'==================
ReDim Modality(InstanceCounter - 1) 'store each exam modality type
ReDim StudyInstanceLocations(InstanceCounter - 1)

RECISTSumT = 0.01
RECISTSumNT = 0.01
ReDim RECISTValsT(InstanceCounter - 1) 'Dimension array for #studies
ReDim RECISTValsNT(InstanceCounter - 1) 'Dimension array for #studies
ArrayPos = 0
iVal = 0
TTrack = 0
NTTrack = 0
NLTrack = 0
EOCE = 0

'==============================================================================
'Finding EOCE, and study header locations, exam modality for printing purposes
'==============================================================================
With ws
    If InstanceCounter = 1 Then
        EOCE = LastRow
    Else
        For StartNumber = 3 To LastRow
            'If InStr(1, ActiveSheet.Cells(StartNumber, StdDescpLoc), "STUDY INSTANCE UID:") > 0 Or IsEmpty(ActiveSheet.Cells(StartNumber, RECISTDiaLoc)) = True Then
            If IsEmpty(.Cells(StartNumber, RECISTDiaLoc)) = True Then
                EOCE = StartNumber
                Exit For 'Breaks at first instance of STUDY INSTANCE UID other than the one for the current exam
            End If
        Next StartNumber
    End If


    'Find study header locations
    For StartNumber = 2 To LastRow
        CellText = .Cells(StartNumber, StdDescpLoc)
        If InStr(1, CellText, "STUDY INSTANCE UID") > 0 Then
            StudyInstanceLocations(ArrayPos) = StartNumber
            ArrayPos = ArrayPos + 1
        End If
    Next StartNumber

    '=============================
    'Find and store exam modality
    '=============================
    For StartNumber = 0 To InstanceCounter - 1
        CellText = .Cells(StudyInstanceLocations(StartNumber) + 1, StdDescpLoc + 1)
            If InStr(1, CellText, "CT") > 0 Then
                Modality(StartNumber) = "CT"
            ElseIf InStr(1, CellText, "MR") > 0 Then
                Modality(StartNumber) = "MR"
            End If
    Next StartNumber
End With
    
'==================================================================
'This section runs no matter how many study instances (Atleast 1)
'==================================================================
'NOTE: ArrayPos < 1 is used for TTrack, NTTRack, and NLTrack because we only care to count the number of these lesions in the CURRENT exam - the reason for this is to print correct order in RECIST worksheet if the lesions are numbered
ArrayPos = 0 'Reset the ArrayPos (DO NOT REMOVE THIS LINE)
For StartNumber = 3 To LastRow 'Note, goes 1 further index so that the blank after the last RECIST diameter is registered, casuing the sum to be inserted into the array
    CellText = ws.Cells(StartNumber, TargetLoc) 'store the text to check if Target or Non-Target
    If StartNumber = LastRow Then
        If (InStr(1, CellText, "Target") > 0) And (InStr(1, CellText, "Non") = 0) Then 'Target measurement
            If ArrayPos < 1 Then
                TTrack = TTrack + 1 'get number of targets in most recent exam
            End If
            RECISTSumT = RECISTSumT + ws.Cells(StartNumber, RECISTDiaLoc).value
        ElseIf (InStr(1, CellText, "Non-Target") > 0) Then 'Non-Target measurement
            If ArrayPos < 1 Then
                NTTrack = NTTrack + 1 'get number of nontargets in most recent exam
            End If
            RECISTSumNT = RECISTSumNT + ws.Cells(StartNumber, RECISTDiaLoc).value
        ElseIf InStr(1, ws.Cells(StartNumber, DescripLoc), "New Lesion") > 0 Then
            If ArrayPos < 1 Then
                NLTrack = NLTrack + 1
            End If
        End If
        'Push the last set of RECIST sums to arrays
        RECISTValsT(ArrayPos) = Round(RECISTSumT, 1) 'STORE, StConvert to cm, Round to 1 decimal place
        RECISTValsNT(ArrayPos) = Round(RECISTSumNT, 1) 'STORE, StConvert to cm, Round to 1 decimal place
        RECISTSumT = 0 'reset sum for next study
        RECISTSumNT = 0 'reset sum for next study
        
    ElseIf (TypeName(ws.Cells(StartNumber, RECISTDiaLoc).value) = "Double") Then 'If number is present in col TargetLoc, add it to a running sum (RECISTSumT or RECISTSumNT)
        'We are not at the last row
        If (InStr(1, CellText, "Target") > 0) And (InStr(1, CellText, "Non") = 0) Then 'Target measurement
            If ArrayPos < 1 Then
                TTrack = TTrack + 1 'get number of targets in most recent exam
            End If
            RECISTSumT = RECISTSumT + ws.Cells(StartNumber, RECISTDiaLoc).value
        ElseIf (InStr(1, CellText, "Non-Target") > 0) Then 'Non-Target measurement
            If ArrayPos < 1 Then
                NTTrack = NTTrack + 1 'get number of nontargets in most recent exam
            End If
            RECISTSumNT = RECISTSumNT + ws.Cells(StartNumber, RECISTDiaLoc).value
        ElseIf InStr(1, ws.Cells(StartNumber, DescripLoc), "New Lesion") > 0 Then
            If ArrayPos < 1 Then
                NLTrack = NLTrack + 1
            End If
        End If
        
     ElseIf InStr(1, ws.Cells(StartNumber, StdDescpLoc), "STUDY INSTANCE UID:") > 0 Then
        'Push all vals to arrays and reset the running sums bc we hit a new study header
        RECISTValsT(ArrayPos) = Round(RECISTSumT, 1) 'STORE, StConvert to cm, Round to 1 decimal place
        RECISTValsNT(ArrayPos) = Round(RECISTSumNT, 1) 'STORE, StConvert to cm, Round to 1 decimal place
        ArrayPos = ArrayPos + 1 'Move to next loc in arrays
        RECISTSumT = 0 'reset sum for next study
        RECISTSumNT = 0 'reset sum for next study
    End If
Next StartNumber

RECISTBaselineT = RECISTValsT(InstanceCounter - 1) 'Baseline RECIST Score is stored
RECISTBaselineNT = RECISTValsNT(InstanceCounter - 1) 'Baseline RECIST Score is stored

'Find last column, after which new headers will be typed in.
LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

'======================================================
'Sort the data so that Targets come before Non-Targets
'======================================================
For StartNumber = 0 To InstanceCounter - 1
    If StartNumber = InstanceCounter - 1 Then
        With ws
            Set rng1 = .Range(.Cells(StudyInstanceLocations(StartNumber) + 1, TargetLoc), .Cells(LastRow, TargetLoc))
            Set rng2 = .Range(.Cells(StudyInstanceLocations(StartNumber) + 1, StdDescpLoc), .Cells(LastRow, LastCol))
        End With
        
        ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.Add Key:=rng1, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets(ws.Name).Sort
            .SetRange rng2
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        With ws
            Set rng1 = .Range(.Cells(StudyInstanceLocations(StartNumber) + 1, TargetLoc), .Cells(StudyInstanceLocations(StartNumber + 1) - 1, TargetLoc))
            Set rng2 = .Range(.Cells(StudyInstanceLocations(StartNumber) + 1, StdDescpLoc), .Cells(StudyInstanceLocations(StartNumber + 1) - 1, LastCol))
        End With
        
        ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.Add Key:=rng1, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            
        With ActiveWorkbook.Worksheets(ws.Name).Sort
            .SetRange rng2
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
Next StartNumber


'==================================================================
'This section runs IF THERE IS ATLEAST 2 STUDIES (Percent changes)
'==================================================================
If InstanceCounter > 1 Then
    'Now need to calculate the RECIST Percentages w.r.t BASELINE
    ReDim RECISTPercentT(InstanceCounter - 1) 'InstanceCounter-1 percent differences (no percent for baseline), leaves an extra space at end so no err when printing vals to sheet
    ReDim RECISTPercentNT(InstanceCounter - 1)

    For StartNumber = 0 To InstanceCounter - 2
        If RECISTBaselineT > 0 Then
            RECISTPercentT(StartNumber) = Round(100 * (RECISTValsT(StartNumber) - RECISTBaselineT) / RECISTBaselineT, 0) 'gives a PERCENTAGE to 1 decimal place
        End If
        If RECISTBaselineNT > 0 Then
            RECISTPercentNT(StartNumber) = Round(100 * (RECISTValsNT(StartNumber) - RECISTBaselineNT) / RECISTBaselineNT, 0) 'gives a PERCENTAGE to 1 decimal place
        End If
    Next StartNumber
    
    '==================================
    '5 Fields for CCR RECIST Worksheet
    '==================================
    CTLS = RECISTValsT(0)
    TLBS = RECISTBaselineT
    '3 others are calculates, BestResponse already stored
    
    '============================================
    '   Best response and percent changes ***BEST RESPONSE IS INCORRECT, NEED TO CONSIDER NEW LESION
    '=============================================
    'Best response is exam after baseline with smallest RECIST sum and NO new lesions
    iVal = RECISTValsT(0)
    For StartNumber = 1 To (InstanceCounter - 2) '-2 to exclude the baseline
        If RECISTValsT(StartNumber) < iVal Then
            iVal = RECISTValsT(StartNumber)
        End If
    Next StartNumber
    BestResponse = iVal 'prior best response stored

    
    '====================================================================================
    'Identify the patient response and categorize
    'Target lesions classified based on RECIST Guidelines and classify as PD, PR, SD, CR
    '*********These calculations need to be corrected*************
    '====================================================================================
    'NEW RESPONSE CALC
    If (100 * ((RECISTValsT(0) - BestResponse) / BestResponse)) > 20 And (RECISTValsT(0) > (BestReponse + 0.5)) Then
        'PD
        ResponseT = "PD"
    ElseIf RECISTPercentT(0) < -30 Then
        'PR
        ResponseT = "PR"
    ElseIf RECISTValsT(0) = 0 Then
        'CR
        ResponseT = "CR"
    Else
        ResponseT = "SD"
    End If
    
    'Non-Target lesion classification depends on team
    If NTTrack > 0 Then
        If (RECISTValsNT(0) > RECISTBaselineNT) Then
            'PD
            ResponseNT = "PD"
        ElseIf RECISTValsNT(0) = 0 Then
            'CR, no non-targets
            ResponseNT = "CR"
        Else
            'SD
            ResponseNT = "SD"
        End If
    ElseIf NTTrack = 0 Then
        ResponseNT = "-"
    End If
        
    
    'Overall response determined using the table
    If ResponseNT = "-" Then
        OverallResponse = ResponseT
    Else
        If ResponseT = "PD" Or ResponseNT = "PD" Then
            'Overall PD
            OverallResponse = "PD"
        ElseIf ResponseT = "CR" Then
            If ResponseNT = "CR" Then
                OverallResponse = "CR"
            ElseIf ResponseNT = "SD" Then
                OverallResponse = "PR"
            End If
        ElseIf ResponseT = "PR" Then
            OverallResponse = "PR"
        ElseIf ResponseT = "SD" Then
            OverallResponse = "SD"
        End If
    End If
    
    '=======================================================
    '*******These calculations need to be corrected********
    '=======================================================
End If

'==========================================================================
'Now need to print the data to the spreadsheet for reading and compilation
'==========================================================================
'Now need to print in to the cells the new header titles
With ws
    
    .Cells(1, LastCol + 1) = "Target RECIST Sum (cm)"
    .Cells(1, LastCol + 2) = "Target RECIST Percent Change (%)"
    .Cells(1, LastCol + 3) = "Non-Target RECIST Sum (cm)"
    .Cells(1, LastCol + 4) = "Non-Target RECIST Percent Change (%)"
    .Cells(1, LastCol + 5) = "Best Response Sum (cm)"
    .Cells(1, LastCol + 6) = "Target Response"
    .Cells(1, LastCol + 7) = "Non-Target Response"
    .Cells(1, LastCol + 8) = "Overall Response"
    .Cells(1, LastCol + 9) = "Exam Type"
    
    .Cells(1, LastCol + 10) = "Patient Name"
    .Cells(1, LastCol + 11) = "MRN"
    .Cells(1, LastCol + 12) = "Protocol #"
    
    
    If InstanceCounter > 1 Then
        For StartNumber = 0 To InstanceCounter - 1
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 1) = RECISTValsT(StartNumber)
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 2) = RECISTPercentT(StartNumber)
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 3) = RECISTValsNT(StartNumber)
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 4) = RECISTPercentNT(StartNumber)
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 9) = Modality(StartNumber)
        Next StartNumber
        
        .Cells(2, LastCol + 5) = BestResponse 'Prints out the best response, only if there is more than 1 exam
        .Cells(2, LastCol + 6) = ResponseT
        .Cells(2, LastCol + 7) = ResponseNT
        .Cells(2, LastCol + 8) = OverallResponse
        
    Else
        For StartNumber = 0 To InstanceCounter - 1
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 1) = RECISTValsT(StartNumber)
            .Cells(StudyInstanceLocations(StartNumber), LastCol + 2) = RECISTValsNT(StartNumber)
        Next StartNumber
    End If
    
    'Not yet Implemented, MRN and Protocol# require proper file naming
    .Cells(2, LastCol + 10) = ws.Cells(3, PatNameLoc).value
    .Cells(2, LastCol + 11) = MedicalRecordNumber
    .Cells(2, LastCol + 12) = ProtocolNumber
    
End With

'Format these new headers
With Range(Cells(1, LastCol + 1), Cells(1, LastCol + 12))
    .Interior.Color = RGB(220, 220, 220)
    .Font.Size = 8.5
    .Font.Name = "Tahoma"
    .Font.Bold = False
    .Borders.LineStyle = XlLineStyle.xlContinuous
End With

End Sub
