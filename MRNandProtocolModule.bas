Attribute VB_Name = "MRNandProtocolModule"
Sub MRNandProtocol(fn As String)
'fn is the FileName as a string, obtained from either Batch or Single Import subprocedures
'Structure of filename is MRN#xxxxxxx_xx-?-xxxx.xml (x is number, ? is
'This function will extract them as two seperate strings for DB management
'Global vars ProtocolNumber, MedicalRecordNumber are used for storage

'Dim tempStr As String
Dim openPos, closePos, midpPos As Integer
    
    openPos = InStr(fn, "#")
    closePos = InStr(fn, ".")
    midpos = InStr(fn, "_")
    
    MedicalRecordNumber = Mid(fn, openPos + 1, midpos - openPos - 1)
    ProtocolNumber = Mid(fn, midpos + 1, closePos - midpos - 1)
End Sub

Sub test()
    MRNandProtocol ("MRN#123456_15-C-0160.xlsx")
    MsgBox MedicalRecordNumber
    MsgBox ProtocolNumber
End Sub
