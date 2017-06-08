Attribute VB_Name = "FolderSetup"
Sub SetupFolders()
'This sub will set up the saving/organizational system of the interface
'Overview: Master Folder --> Dr's Folder --> Specific Date/TimeFolder --> PDF, Word, ExcelSheets Output folders
'Labmatrix will be given a seperate folder: Master Folder --> Labmatrix --> Dr's Folder --> Date/Time of Export


'**************
'Might change this so that user is asked for the FOLDER they want, and that is created within master-->outputs-->foldername

'Variables:
Dim InterfaceParent As String
Dim strDate, strTime As String
Dim Username As String
strDate = Format(Date, "mm-dd-yyyy")
strTime = Format(Time, "hh.nn AM/PM")

Username = UserInterface.LastName.value
If Username = "" Then
    'In case no name entered
    Username = "Default"
End If

'====================================
'Create the 'Master Interface Folder'
'====================================
InterfaceParent = ThisWorkbook.path
'InterfaceParent = CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder.Path
If Len(Dir(InterfaceParent & "\Master Folder", vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder"
End If

'==================
'Create 'Outputs'
'==================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs", vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs"
End If

'=======================
'Create Doctor's folder
'=======================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username), vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username)
End If

''=======================
'Create todays date
'========================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate, vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate
End If

'====================
'Create todays time
'====================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime, vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime
End If

'====================
'Create WordDocs
'====================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "WordDocs", vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "WordDocs"
End If

'=======================
'Create Labmatrix folder
'=======================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Labmatrix", vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Labmatrix"
End If

'=======================
'Create OutputSheets folder
'=======================
If Len(Dir(InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Output Sheets", vbDirectory)) = 0 Then
   MkDir InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Output Sheets"
End If

OutputSheetsLoc = InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Output Sheets"
OutputLoc = InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime
LabmatrixLoc = InterfaceParent & "\Master Folder\Outputs" & "\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "Labmatrix\"
WordDocLoc = InterfaceParent & "\Master Folder\Outputs\" & UCase(Username) & "\" & strDate & "\" & strTime & "\" & "WordDocs\"


End Sub
Function getParentFolder2(ByVal strFolder0)
  Dim strFolder
  strFolder = Left(strFolder0, InStrRev(strFolder0, "\") - 1)
  getParentFolder2 = strFolder
  'getParentFolder2 = Left(strFolder, InStrRev(strFolder, "\") - 1)
End Function
