Attribute VB_Name = "DatabaseUploader"
Sub DBUploader()

'====================================
'Database login for upload
'===================================
Dim Username, Password, StrCommmand, FileName, UploadFilePath As String

Username = LoginPage.Username.value
Password = LoginPage.Password.value

ChDir LabmatrixLoc
UploadFilePath = Application.GetOpenFilename(("Excel Files (*.xl*), *.xl*"), 1, "Select File")
    
StrCommand = "java -cp" & ThisWorkbook.path & "\RadiologyImportClient.jar ImportArgs " & Username & " " & Password & " " & """" & UploadFilePath & """"
Shell StrCommand, 1 ' Change to 0 makes cmd prompt not appear?

End Sub
