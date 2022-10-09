Attribute VB_Name = "Point2DFactory"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Function MakePoint2D(ByRef xCoordinate As Double, ByRef yCoordinate As Double) As Point2D

    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = xCoordinate
        .y = yCoordinate
    End With
    
    Set MakePoint2D = point
    
End Function
