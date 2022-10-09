Attribute VB_Name = "Vector2DFactory"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

'@Description "Return Vector2D object given the end coordinates of a vector."
Public Function MakeVector(ByRef u As Double, ByRef v As Double) As Vector2D
Attribute MakeVector.VB_Description = "Return Vector2D object given the end coordinates of a vector."
    Dim Vector As Vector2D
    Set Vector = New Vector2D
    With Vector
        .u = u
        .v = v
    End With
    
    Set MakeVector = Vector
End Function

'@Description "Return Vector2D object between two Point2D objects."
Public Function MakeVectorBetween(ByRef point1 As Point2D, ByRef point2 As Point2D) As Vector2D
Attribute MakeVectorBetween.VB_Description = "Return Vector2D object between two Point2D objects."
    Set MakeVectorBetween = point2.Subtract(point1)
End Function

'@Description "Return Vector2D object of a unit length given the end coordinates of a vector."
Public Function MakeUnitVector(ByRef u As Double, ByRef v As Double) As Vector2D
Attribute MakeUnitVector.VB_Description = "Return Vector2D object of a unit length given the end coordinates of a vector."
    Dim Vector As Vector2D
    Set Vector = New Vector2D
    With Vector
        .u = u
        .v = v
    End With
    
    Set MakeUnitVector = Vector.Normalized
End Function

'@Description "Return Vector2D object of a unit length between two Point2D objects."
Public Function MakeUnitVectorBetween(ByRef point1 As Point2D, ByRef point2 As Point2D) As Vector2D
Attribute MakeUnitVectorBetween.VB_Description = "Return Vector2D object of a unit length between two Point2D objects."
    Set MakeUnitVectorBetween = MakeVectorBetween(point1, point2).Normalized
End Function
