Attribute VB_Name = "RECISTMainModule"
'=================================================================
'Written By Nikhil Goyal
'Version 5.1, 8/30/2016
'last edit 8/15/2016
'=================================================================

Sub RECIST() 'Application.DisplayAlerts = True 'http://www.excely.com/excel-vba/disable-alert-warning-messages.shtml

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    .DisplayStatusBar = False
End With

'VARIABLES:
'going through all worksheets
Dim ws As Worksheet
Dim SkipFlag, wsFlag As Integer
Const dots As String = "..........."
Dim StartTime, SecondsElapsed As Double 'For timing the script
Dim HLink As String

StartTime = Timer

Call SetupFolders
Call DeleteAllButMain   'clear before running

'PROGRESS BAR:
Set MyProgressBar = New ProgressBar
With MyProgressBar
    .Title = "Script Progress" 'Set the Title
    .ExcelStatusBar = True 'Set this to true if you want to update Excel's Status Bar Also
    .StartColour = rgbRed 'Set the colour of the bar in the Beginning
    .EndColour = rgbGreen ' Set the colour of the bar at the end
End With
MyProgressBar.TotalActions = 100
MyProgressBar.ShowBar


'==========================================
'User Interface Code, IMPORTATION
'==========================================
With MyProgressBar
    .ActionNumber = 5
    .StatusMessage = "Importing Files"
End With

Call FileSelectAndImport

'==========================================================================
'Add a sheet called "Output" where the data will be compiled for waterfall
'==========================================================================
With ThisWorkbook
    If .Sheets.count > 1 Then
        .Sheets.Add After:=Worksheets("Main")
        .Sheets.Item(2).Name = "Output"
        .Worksheets("Output").Cells(1, "A") = "File"
        .Worksheets("Output").Cells(1, "B") = "Patient Name"
        .Worksheets("Output").Cells(1, "C") = "MRN"
        .Worksheets("Output").Cells(1, "D") = "Protocol #"
        .Worksheets("Output").Cells(1, "E") = "Current Target Lesion Sum % Change from Baseline"
        .Worksheets("Output").Cells(1, "F") = "Current Target Lesion Sum % Change from Best Response"
        .Worksheets("Output").Cells(1, "G") = "Best Response % Change from Baseline"
        .Worksheets("Output").Cells(1, "H") = "Current Non-Target Lesion Sum % Change from Baseline"
        
        With Range("A1", "H1")
            .Interior.Color = RGB(220, 220, 220)
            .Font.Size = 8.5
            .Font.Name = "Tahoma"
            .Font.Bold = False
            .Borders.LineStyle = XlLineStyle.xlContinuous
        End With
    End If
End With

'============================
'Iterate through the sheets
'============================
With MyProgressBar
    .ActionNumber = 25
    .StatusMessage = "Files imported. Now cleaning and performing RECIST calculations"
End With

wsFlag = 2 'indicates the row to print on in the 'Output' sheet
For Each ws In ThisWorkbook.Worksheets 'Run on each sheet
    If ws.Name <> "Main" And ws.Name <> "Output" And ws.Name <> "Combined" Then

        
        '===========================
        'get protocol# & medrecord#
        '===========================
        Call MRNandProtocol(ws.Name)
        
        SkipFlag = FindCols(ws, StdDescpLoc, PatNameLoc, FllwUpLoc, NameLoc, ToolLoc, _
                    DescripLoc, TargetLoc, SubTypeLoc, SeriesLoc, SliceLoc, RECISTDiaLoc, LongDiaLoc, ShortDiaLoc, _
                    CreatorLoc, LengthLoc, VolumeLoc, HUMeanLoc, PODLoc)
        
        If Len(ws.Name) > 26 Then
            SkipFlag = 100
        End If
        
        If SkipFlag <> 0 Then
            'Missing column
            If SkipFlag <> 100 Then
                'MsgBox ws.Name & " was skipped and removed because it is improperly formatted: a mandatory field is missing"
                With Report.ReportBox
                    .Text = .Text & ws.Name & dots & "Import Failed" & dots & "Bookmark table missing a category" & vbNewLine
                End With
            ElseIf SkipFlag = 100 Then
                'MsgBox ws.Name & " was skipped and removed because it is improperly named. Please use the naming convention 'MRN#xxxxxxx_xx-*-xxx' where x represents a number, and * is a letter"
                With Report.ReportBox
                    .Text = .Text & ws.Name & dots & "Import Failed" & dots & "Incorrect file name format" & vbNewLine
                End With
            End If
        
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
            ws.Delete

        ElseIf SkipFlag = 0 Then
            With Report.ReportBox
                .Text = .Text & ws.Name & dots & "Import Sucess" & dots & "None" & vbNewLine
            End With
            '==============================================
            'Now clean the sheet according to RECIST needs, all extraneous info
                '(exams prior to baseline, lesions NOT labeled as Target/Non-Target, and exams after baseline without targets/non-targets) removed
            '==============================================
            Call CleanSheet(ws)
            
            '========================================
            'Perform the RECIST Calcs, use a funtion
            '========================================
            Call RECISTCalc(ws)
            
            '========================================
            'Print to the CCR RECIST WorkSheet
            '========================================
            If UserInterface.WordGen = True Then
                Call RECISTWordExport(ws)
            End If
            
            '===========================================
            'Compile data for later work
            '===========================================
            With Worksheets("Output") 'ONLY INCLUDES THOSE WITH FOLLOWUP EXAMS
                If InstanceCounter > 1 And ThisWorkbook.Sheets.count > 1 Then
                    .Cells(wsFlag, "A") = ws.Name
                    .Cells(wsFlag, "B") = ws.Cells(3, PatNameLoc)
                    '.Cells(wsFlag, "C") = MedicalRecordNumber
                    HLink = "https://radlite.cc.nih.gov/portal?password_encrypted&hide_top=all&patient_id=" & MedicalRecordNumber
                    .Hyperlinks.Add Anchor:=.Cells(wsFlag, "C"), Address:=HLink _
                        , TextToDisplay:=MedicalRecordNumber
                    .Cells(wsFlag, "D") = ProtocolNumber
                    .Cells(wsFlag, "E") = RECISTPercentT(0)
                    .Cells(wsFlag, "F") = Round(100 * ((RECISTValsT(0) - BestResponse) / BestResponse), 0)
                    .Cells(wsFlag, "G") = Round(100 * ((BestResponse - RECISTBaselineT) / RECISTBaselineT), 0)
                    .Cells(wsFlag, "H") = RECISTPercentNT(0)
                    wsFlag = wsFlag + 1
                End If
            End With
        End If
    End If
    ws.Cells.EntireColumn.AutoFit 'For readability
Next

'============================================================================================
'Make 'Outputs' look pretty, and organize from lowest to highest best response from baseline
'============================================================================================
If WorksheetExists("Output") = True Then
    Sheets("Output").Cells.EntireColumn.AutoFit 'For readability
    Worksheets("Output").Range("G2").CurrentRegion.Sort Key1:=Worksheets("Output").Range("G2"), Order1:=xlDescending, Header:=xlGuess 'Sort for waterfall plot
End If



'=============================================================
'Check if word is still open, if it is, CLOSE the application
'=============================================================


'========================================
'Create PDF RECIST Documents
'========================================
If UserInterface.RECISTPDFOption = True Then
    With MyProgressBar
        .ActionNumber = 75
        .StatusMessage = "Sheets created, now generating PDF Files"
    End With
    Call PDF
Else
    With MyProgressBar
        .ActionNumber = 100
        .StatusMessage = "Sheets created."
    End With
End If
'Call PDF

'========================================
'If specified by user, compile data
'========================================
If UserInterface.CompileSheets = True Then
    Call CombineSheets
End If

'============================================
'Autosave the 'Output' and 'Compiled' Sheets
'============================================
If UserInterface.SaveOutputSheets = True Then
    Call SaveOutputSheets
End If

'==========================================
'Create files for Labmatrix
'==========================================
If UserInterface.RECISTLabmatrixOutputOption = True Then
    Call LMWkbkFormatter(Workbooks(InterfaceVersion))
End If
    
'========================================
'Clean the document, deleting sheets
'========================================
If UserInterface.DiscardSheets = True Then
    Call DeleteSheets
End If

With MyProgressBar
    .ActionNumber = 100
    .StatusMessage = "Done!"
End With
MyProgressBar.Complete 1

If UserInterface.ShowReport = True Then
    Report.Show
End If

SecondsElapsed = Round(Timer - StartTime, 0)

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
    .DisplayStatusBar = True
End With


End Sub
