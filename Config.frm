VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Config 
   Caption         =   "Interface Configuration"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14325
   OleObjectBlob   =   "Config.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeBTP_Click()
'MsgBox "Please locate the folder containing bookmark tables"
Cells(1, "A").value = GetFolder("C:\") & "\"
BookmarkTablePath = ThisWorkbook.Sheets("Main").Cells(1, 1).value
End Sub
Private Sub ChangeLOP_Click()
Cells(5, "A").value = GetFolder("C:\") & "\"
LabmatrixOutputPath = ThisWorkbook.Sheets("Main").Cells(5, 1).value
End Sub

Private Sub ChangeOP_Click()
'MsgBox "Please locate the folder where ouputs will save to"
Cells(4, "A").value = GetFolder("C:\") & "\"
OutputPath = ThisWorkbook.Sheets("Main").Cells(4, 1).value
End Sub

Private Sub ChangePDFCP_Click()
Cells(3, "A").value = GetFolder("C:\") & "\"
RECISTFormPath = ThisWorkbook.Sheets("Main").Cells(3, 1).value
PDFCPath = ThisWorkbook.Sheets("Main").Cells(3, 1).value
End Sub
Private Sub ChangeRFP_Click()
'MsgBox "Please locate the RECIST and PDFC worksheets"
Cells(3, "A").value = GetFolder("C:\") & "\"
RECISTFormPath = ThisWorkbook.Sheets("Main").Cells(3, 1).value
PDFCPath = ThisWorkbook.Sheets("Main").Cells(3, 1).value
End Sub


Private Sub ResetConfig_Click()
Dim reset As Integer
        reset = MsgBox("Do you want to reset configuration?", vbYesNo + vbQuestion, "Empty Sheet")
        If reset = vbYes Then
            BookmarkTablePath = "/"
            OutputPath = "/"
            RECISTFormPath = "/"
            PDFCPath = "/'"
            OutputSavePath = "/"
        'Save Prefs
        'Workbooks(ThisWorkbook.Name).Save
        'Workbooks.Open (Workbooks(ThisWorkbook.Name).Path & "\" & ThisWorkbook.Name)
        End If
End Sub
Private Sub UserForm_Activate()

'Start Userform Centered inside Excel Screen (for dual monitors)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
  
BookmarkTablePath = ThisWorkbook.Sheets("Main").Cells(1, 1).value
OutputPath = WordDocLoc
RECISTFormPath = Application.ActiveWorkbook.path
PDFCPath = Application.ActiveWorkbook.path
OutputSavePath = WordDocLoc
LabmatrixOutputPath = LabmatrixLoc

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Workbooks(ThisWorkbook.Name).Save 'save when config closes
End Sub
