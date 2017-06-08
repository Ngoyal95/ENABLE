VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserInterface 
   Caption         =   "Semi-Automation Interface"
   ClientHeight    =   12525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   OleObjectBlob   =   "UserInterface.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseUI_Click()
    Call DeleteAllButMain 'Clean before saving
    Workbooks(ThisWorkbook.Name).Save
    Unload Me
    'Application.Quit
    Excel.Application.Quit
End Sub

Private Sub CommandButton1_Click()
    Call DeleteAllButMain
    'Call DeleteSheets
End Sub

Private Sub ConfigUI_Click()
    'Call ConfigProgram
    Config.Show
End Sub

Private Sub Help_Click()
    HelpPage.Show
End Sub


Private Sub OtherRun_Click()
    Me.Hide
    'Application.Visible = False 'Makes excel invisible.
    Call ApoloMain
    'Application.Visible = True 'Makes excel visible.
    Me.Show
End Sub
Private Sub OtherClearSpreadsheets_Click()
    Call DeleteSheets
End Sub
Private Sub GenRECISTReport_Click()
    Me.Hide
    Call RECIST
    Me.Show
End Sub

Private Sub HideExcel_Click()
    Application.Visible = False 'Makes excel invisible.
End Sub

Private Sub RECISTSaveOutputs_Click()
Dim SaveName, flpt As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If WorksheetExists("Combined") = False Then
        MsgBox "Outputs and/or Combined sheets not generated! Run the program."
    Else
        SaveName = InputBox("Enter name to save file as")
        flpt = GetFolder(ActiveWorkbook.Sheets("Main").Cells(4, "A").value) & "\" & SaveName & ".xlsx"
        Sheets(Array("Combined")).Move
        ActiveWorkbook.SaveAs FileName:=flpt, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
Private Sub SaveOutouts_Click()
'Save the "Output" and "Combined" Sheets into a new workbook, does NOT delete them
Dim SaveName, flpt As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If WorksheetExists("Combined") = False Then
        MsgBox "Outputs and/or Combined sheets not generated! Run the program."
    Else
        SaveName = InputBox("Enter name to save file as")
        flpt = GetFolder(ActiveWorkbook.Sheets("Main").Cells(4, "A").value) & "\" & SaveName & ".xlsx"
        Sheets(Array("Combined", "Output")).Move
        ActiveWorkbook.SaveAs FileName:=flpt, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub ShowExcel_Click()
    Application.Visible = True 'Makes excel visible.
End Sub

Private Sub ShowFolders_Click()
    Call Shell("explorer.exe" & " " & OutputLoc, vbNormalFocus)
End Sub

Private Sub UploadtoDatabases_Click()
    LoginPage.Show
End Sub

Private Sub UserForm_Initialize()
    'Start Userform Centered inside Excel Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
  
    UserInterface.BackColor = RGB(224, 255, 255)
    Me.Label1.BackColor = RGB(224, 255, 255)
    Me.Label5.BackColor = RGB(224, 255, 255)
    Me.Label6.BackColor = RGB(224, 255, 255)
    Me.DiscardSheets = True 'automatically choose to discard sheets
    'Me.ShowReport = True 'automatically fill checkbox to show the report
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  Application.Visible = True
  If CloseMode = vbFormControlMenu Then
    'Cancel = True
    'MsgBox "Please use the Close Form button!"
    'Cancel = True
    'Call DeleteSheets 'Clean before saving
  End If
End Sub



