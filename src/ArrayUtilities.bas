Attribute VB_Name = "ArrayUtilities"
'@Folder("StructuralAnalysis.Utilities")
Option Explicit

Public Function ArrayLength(ByRef arr() As Double) As Long
    If IsEmpty(arr) Then
        ArrayLength = 0
    Else
        ArrayLength = UBound(arr) - LBound(arr) + 1
    End If
End Function
