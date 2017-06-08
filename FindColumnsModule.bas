Attribute VB_Name = "FindColumnsModule"
Function FindCols(ws As Worksheet, StdDescpLoc, PatNameLoc, FllwUpLoc, NameLoc, ToolLoc, _
DescripLoc, TargetLoc, SubTypeLoc, SeriesLoc, SliceLoc, RECISTDiaLoc, LongDiaLoc, ShortDiaLoc, _
CreatorLoc, LengthLoc, VolumeLoc, HUMeanLoc, PODLoc As Integer) As Integer

'Function returns 0 if all mandatory fields present, returns >0 otherwise

'VARIABLES:
Dim rng As Range
Dim row As Range
Dim cell As Range
Dim MissingFlag As Integer 'Used to escape the For Each


'Initialize all variables to zero:
StdDescpLoc = 0
PatNameLoc = 0
FllwUpLoc = 0
NameLoc = 0
ToolLoc = 0
DescripLoc = 0
TargetLoc = 0
SubTypeLoc = 0
SeriesLoc = 0
SliceLoc = 0
RECISTDiaLoc = 0
LongDiaLoc = 0
ShortDiaLoc = 0
CreatorLoc = 0
LengthLoc = 0
VolumeLoc = 0
HUMeanLoc = 0
PODLoc = 0

MissingFlag = 9
Set rng = ws.Range("A1").EntireRow 'range of categories

For Each cell In rng.Rows.Cells
    Select Case cell
        Case "Study Description"
        'mandatory field
        StdDescpLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Patient Name"
        'mandatory field
        PatNameLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Follow-Up"
        'Mandatory field
        FllwUpLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Name"
        'Not mandatory
        NameLoc = cell.Column
        
        
        Case "Tool"
        'NOT mandatory
        ToolLoc = cell.Column
        
        
        Case "Description"
        'MANDATORY
        DescripLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Target"
        'MANDATORY
        TargetLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Sub-Type"
        'NOT MANDATORY
        SubTypeLoc = cell.Column
        
        
        Case "Series"
        'MANDATORY
        SeriesLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Slice#"
        'MANDATORY
        SliceLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "RECIST Diameter ( mm )"
        'MANDATORY
        RECISTDiaLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Long Diameter ( mm )"
        'NOT MANDATORY
        LongDiaLoc = cell.Column
        
        
        Case "Short Diameter ( mm )"
        'NOT MANDATORY
        ShortDiaLoc = cell.Column
        
        Case "Creator"
        'MANDATORY
        CreatorLoc = cell.Column
        MissingFlag = MissingFlag - 1
        
        Case "Length ( mm )"
        LengthLoc = cell.Column
        
        Case "Volume ( mm³ )"
        VolumeLoc = cell.Column
        
        Case "HU Mean(HU)"
        HUMeanLoc = cell.Column
        
        Case "Product of Diameters ( mm² )"
        PODLoc = cell.Column
        
    End Select
Next cell

FindCols = MissingFlag

End Function
