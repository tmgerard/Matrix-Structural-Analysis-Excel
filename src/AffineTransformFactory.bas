Attribute VB_Name = "AffineTransformFactory"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Function MakeScaleTransformAboutPoint(ByRef ScaleX As Double, _
                                             ByRef ScaleY As Double, _
                                             ByRef center As Point2D) As AffineTransform

    Dim scaled As AffineTransform
    Set scaled = New AffineTransform
    With scaled
        .ScaleX = ScaleX
        .ScaleY = ScaleY
        .translateX = center.x * (1# - ScaleX)
        .translateY = center.y * (1# - ScaleY)
    End With
    
    Set MakeScaleTransformAboutPoint = scaled

End Function

Public Function MakeRotationTransformAboutPoint(ByRef radians As Double, ByRef center As Point2D) As AffineTransform

    Dim cos As Double
    cos = Math.cos(radians)
    
    Dim sin As Double
    sin = Math.sin(radians)
    
    Dim rotatedTransform As AffineTransform
    Set rotatedTransform = New AffineTransform
    With rotatedTransform
        .ScaleX = cos
        .ScaleY = sin
        .translateX = center.x * (1 - cos) + center.y * sin
        .translateY = center.y * (1 - cos) - center.x * sin
        .shearX = -sin
        .shearY = sin
    End With
    
    Set MakeRotationTransformAboutPoint = rotatedTransform

End Function
