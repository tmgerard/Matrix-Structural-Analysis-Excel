Attribute VB_Name = "Segment2DRatio"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Const MIN As Double = 0#
Public Const MID As Double = 0.5
Public Const MAX As Double = 1#

Public Enum SegmentRatioError
    BadValue = 100 + vbObjectError
End Enum

Public Function MakeValidRatio(ByRef value As Double) As Double
    If value < MIN Then
        MakeValidRatio = MIN
    ElseIf value > MAX Then
        MakeValidRatio = MAX
    Else
        MakeValidRatio = value
    End If
End Function

Public Sub EnsureValidRatio(ByRef ratio As Double)
    If ratio < MIN Or ratio > MAX Then
        Err.Raise Number:=SegmentRatioError.BadValue, _
                  source:="Segment2DRatio", _
                  Description:="Expected ratio to be in [0, 1] but was " & ratio
    End If
End Sub

Public Function IsValid(ByRef ratio As Double) As Boolean
    IsValid = Not ratio < MIN And Not ratio > MAX
End Function
