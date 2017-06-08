VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HelpPage 
   Caption         =   "Help"
   ClientHeight    =   11670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14280
   OleObjectBlob   =   "HelpPage.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HelpPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'Start Userform Centered inside Excel Screen (for dual monitors)
Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Me.Hide
   UserInterface.Show
End Sub

