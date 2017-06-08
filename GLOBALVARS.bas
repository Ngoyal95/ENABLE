Attribute VB_Name = "GLOBALVARS"
'================
'Program version
'================
Public Const InterfaceVersion As String = "Version5.1.xlsm"

'=========================
'SHARED PARSER VARIABLES
'=========================
'Finding category locations (what colums) (Accessed by print module)
Public StdDescpLoc, PatNameLoc, FllwUpLoc, NameLoc, ToolLoc, DescripLoc, TargetLoc, _
SubTypeLoc, SeriesLoc, SliceLoc, RECISTDiaLoc, LongDiaLoc, ShortDiaLoc, CreatorLoc, _
LengthLoc, VolumeLoc, HUMeanLoc, PODLoc As Integer

'Push down blank rows
Public s1 As Worksheet
Public tmpR As Range
Public rowcount As Long, colcount As Long, i As Long, j As Long, k As Boolean

'LastRow and last column
Public LastRow, LastCol As Integer

Public StudyInstanceLocations As Variant 'Track where study headers are located

'===================
'RECIST VARIABLES
'===================
'RECIST Cals (accessed by word print module)
Public InstanceCounter, EOCE As Integer 'tracks the number of trails, EOCE used to indicate row where current exam breaks
Public RECISTSumT, RECISTBaselineT, RECISTSumNT, RECISTBaselineNT As Double 'Variable used to calc the RECIST values, located in Col L, CURRENTLY IN mm. Both T (target) and NT (nontarget)
Public StudyInstanceCounter As Integer 'Counts the #of studies present (in the cleaned sheet)
Public CTLS, TLBS As Double

'RECIST RESPONSES
Public BestResponse As Double
Public ResponseT, ResponseNT, OverallResponse As String

'RECIST ARRAYS
Public RECISTValsT, RECISTValsNT As Variant 'Arrays used to store the recist diameters
Public RECISTPercentT, RECISTPercentNT As Variant 'Arrays to store the percentages

'Vars to track number of targets and non-targets (needed for recist sheet printing)
Public TTrack, NTTrack, NLTrack As Integer


'================================================
'Patient protocol# and Medical record# trackers, Modality (CT or MR)
'================================================
Public ProtocolNumber, MedicalRecordNumber As String 'acquired from the filename which is formatted as protocol#-medicalrecord#.xml
Public Modality As Variant 'Tracks if an exam was CT or MR, array since exam type can vary

'========================
'Script Progress tracker
'========================
Public MyProgressBar As ProgressBar

'=====================================
'Filepaths for saving
'=====================================
Public OutputLoc, WordDocLoc, LabmatrixLoc, OutputSheetsLoc As String


