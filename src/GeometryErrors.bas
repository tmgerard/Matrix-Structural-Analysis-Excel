Attribute VB_Name = "GeometryErrors"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Enum GeometryError
    NegativeDimension = 900 + vbObjectError
End Enum

Public Sub RaiseNegativeDimensionError(Optional ByRef message As String = vbNullString, _
    Optional ByRef source As String = vbNullString)
    ' Raise an error for a zero or negative dimension
    Err.Raise Number:=GeometryError.NegativeDimension, _
              source:=source, _
              Description:=message

End Sub

