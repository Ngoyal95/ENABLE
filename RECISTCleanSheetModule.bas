Attribute VB_Name = "RECISTCleanSheetModule"
Sub CleanSheet(ws As Worksheet)
'Cleaning sheet for RECIST Form

'============
'VARIABLES:
'============
'Dim StartNumber, LastRow As Integer 'Startnumber is used in the forloop, LastRow stores the index of the last row filled with data (Col B is used for Bookmark Table)
Dim StartNumber As Integer 'Startnumber is used in the forloop, LastRow stores the index of the last row filled with data (Col B is used for Bookmark Table)
Dim CellText As String 'used to store the cell's string

'Finding headers to be removed
Dim RemoveCount As Integer 'used to go to the new last row
Dim StudyHeaderCount, HeaderLocs() As Integer 'Count # of study headers

'==========
'CODE:
'==========
With ws
    LastRow = .UsedRange.Rows(.UsedRange.Rows.count).row 'Find LastRow
    RemoveCount = 0
    
    If RECISTDiaLoc > 0 Then
        .Cells(1, RECISTDiaLoc).value = "RECIST Diameter (cm)"
    End If
    If LongDiaLoc > 0 Then
        .Cells(1, LongDiaLoc).value = "Long Diameter (cm)"
    End If
    If ShortDiaLoc > 0 Then
        .Cells(1, ShortDiaLoc).value = "Short Diameter (cm)"
    End If
    If LengthLoc > 0 Then
        .Cells(1, LengthLoc).value = "Length (cm)"
    End If
    If VolumeLoc > 0 Then
        .Cells(1, VolumeLoc).value = "Volume (cm³)"
    End If
    If PODLoc > 0 Then
        .Cells(1, PODLoc).value = "Product of Diameters (cm²)"
    End If
    If HUMeanLoc > 0 Then
        .Cells(1, HUMeanLoc).value = "HU Mean (HU)"
    End If
    
    If UserInterface.RECISTClean = True And UserInterface.AAClean = False Then
        '========================
        'remove non-RECIST info
        '========================
        'Now need to parse the rows and remove: exams before baseline, all things not labeled TARGET or NON-TARGET, any exams after baseline without TARGET or NONTARGET
        For StartNumber = LastRow To 2 Step -1 'Run from bottom to top
            CellText = .Cells(StartNumber, TargetLoc)
            If InStr(1, .Cells(StartNumber, StdDescpLoc + 1), "-") > 0 Then
                .Rows(StartNumber).Clear
                RemoveCount = RemoveCount + 1
            ElseIf InStr(1, CellText, "Target") > 0 Or InStr(1, LCase$(.Cells(StartNumber, DescripLoc)), "new lesion") Or IsEmpty(.Cells(StartNumber, StdDescpLoc)) = False Then
                'NOTE: IN THE ABOVE LINE DO NOT REMOVE THE 'Or IsEmpty(.Cells(StartNumber, StdDescpLoc)) = False' OTHERWISE CODE BREAKS
                
                'Convert from mm --> cm
                .Cells(StartNumber, RECISTDiaLoc).value = Round((.Cells(StartNumber, RECISTDiaLoc).value) / 10, 1)
                .Cells(StartNumber, RECISTDiaLoc).NumberFormat = "0.0"
                'Check if Long and Short Diameters are there, convert as well
                If LongDiaLoc > 0 Then
                    .Cells(StartNumber, LongDiaLoc).value = Round((.Cells(StartNumber, LongDiaLoc).value) / 10, 1)
                    .Cells(StartNumber, LongDiaLoc).NumberFormat = "0.0"
                End If
                If ShortDiaLoc > 0 Then
                    .Cells(StartNumber, ShortDiaLoc).value = Round((.Cells(StartNumber, ShortDiaLoc).value) / 10, 1)
                    .Cells(StartNumber, ShortDiaLoc).NumberFormat = "0.0"
                End If
                
                If LengthLoc > 0 Then
                    .Cells(StartNumber, LengthLoc).value = Round((.Cells(StartNumber, LengthLoc).value) / 10, 1)
                    .Cells(StartNumber, LengthLoc).NumberFormat = "0.0"
                End If
                
                If VolumeLoc > 0 Then
                    .Cells(StartNumber, VolumeLoc).value = Round((.Cells(StartNumber, VolumeLoc).value) / 1000, 1)
                    .Cells(StartNumber, VolumeLoc).NumberFormat = "0.0"
                End If
                
                If PODLoc > 0 Then
                    .Cells(StartNumber, PODLoc).value = Round((.Cells(StartNumber, PODLoc).value) / 100, 1)
                    .Cells(StartNumber, PODLoc).NumberFormat = "0.0"
                End If
                
            Else
                'Study Description is blank so it is not an exam header, and it is NOT labeled, delete it:
                .Rows(StartNumber).Clear
                RemoveCount = RemoveCount + 1
            End If
        Next StartNumber
        
    ElseIf UserInterface.AAClean = True And UserInterface.RECISTClean = False Then
        '==================================================================================================
        'Clean non-Apolo info
        'allow the following: 'aa target' + Target, 'aa non-target' + Non-Target, 'aa new lesion' and ????
        '==================================================================================================
        For StartNumber = LastRow To 2 Step -1 'Run from bottom to top
            CellText = .Cells(StartNumber, TargetLoc)
            If InStr(1, .Cells(StartNumber, StdDescpLoc + 1), "-") > 0 Then
                .Rows(StartNumber).Clear
                RemoveCount = RemoveCount + 1
            ElseIf (InStr(1, CellText, "Target") > 0 And InStr(1, LCase$(.Cells(StartNumber, DescripLoc)), "aa target") > 0) _
            Or (InStr(1, LCase$(.Cells(StartNumber, DescripLoc)), "aa new lesion") > 0) Or (IsEmpty(.Cells(StartNumber, StdDescpLoc)) = False) Then
                'Convert from mm --> cm
                .Cells(StartNumber, RECISTDiaLoc).value = Round((.Cells(StartNumber, RECISTDiaLoc).value) / 10, 1)
                .Cells(StartNumber, RECISTDiaLoc).NumberFormat = "0.0"
                'Check if Long and Short Diameters are there, convert as well
                If LongDiaLoc > 0 Then
                    .Cells(StartNumber, LongDiaLoc).value = Round((.Cells(StartNumber, LongDiaLoc).value) / 10, 1)
                    .Cells(StartNumber, LongDiaLoc).NumberFormat = "0.0"
                End If
                If ShortDiaLoc > 0 Then
                    .Cells(StartNumber, ShortDiaLoc).value = Round((.Cells(StartNumber, ShortDiaLoc).value) / 10, 1)
                    .Cells(StartNumber, ShortDiaLoc).NumberFormat = "0.0"
                End If
            Else
                'Study Description is blank so it is not an exam header, and it is NOT labeled, delete it:
                .Rows(StartNumber).Clear
                RemoveCount = RemoveCount + 1
            End If
        Next StartNumber
    End If
    
    Call PushBlanks(ws) 'Push blank rows
    
    '=====================================
    'Find the study headers to be removed
    '=====================================
    LastRow = LastRow - RemoveCount
    StudyHeaderCount = 0
    For StartNumber = 2 To LastRow
        If InStr(1, .Cells(StartNumber, StdDescpLoc), "STUDY INSTANCE UID:") > 0 Then
            If (InStr(1, .Cells(StartNumber + 1, StdDescpLoc), "STUDY INSTANCE UID:") > 0 Or InStr(1, .Cells(StartNumber - 1, StdDescpLoc), "STUDY INSTANCE UID:") > 0) _
            And (IsEmpty(.Cells(StartNumber + 1, StdDescpLoc)) = False Or IsEmpty(.Cells(StartNumber + 1, StdDescpLoc + 1)) = True) Then
                StudyHeaderCount = StudyHeaderCount + 1 'Count number of study headers
                ReDim Preserve HeaderLocs(StudyHeaderCount) ' ReDim
                HeaderLocs(StudyHeaderCount - 1) = StartNumber
            End If
        End If
    Next StartNumber
        
    'Now delete the rows, iterate through all and check if the row# is in the array HeaderLocs, then push blank rows down
    For StartNumber = 0 To StudyHeaderCount - 1
        .Rows(HeaderLocs(StartNumber)).Clear
    Next StartNumber
    
    'Delete last row if needed
    If InStr(1, .Cells(LastRow, StdDescpLoc), "STUDY INSTANCE UID:") > 0 Then
        .Rows(LastRow).Clear
    End If
End With

Call PushBlanks(ws) 'Push blank rows to bottom of sheet


End Sub

