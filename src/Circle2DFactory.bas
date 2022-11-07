Attribute VB_Name = "Circle2DFactory"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Function MakeCircle(ByRef Radius As Double, ByRef center As Point2D) As Circle2D

    Dim circ As Circle2D
    Set circ = New Circle2D
    With circ
        .Radius = Radius
        Set .center = center
    End With
    
    Set MakeCircle = circ

End Function

Public Function MakeCircleFromPoints(ByRef pointA As Point2D, ByRef pointB As Point2D, ByRef pointC As Point2D) As Circle2D
    Dim segment1 As Segment2D
    Set segment1 = New Segment2D
    With segment1
        Set .StartPoint = pointA
        Set .EndPoint = pointB
    End With
    
    Dim segment2 As Segment2D
    Set segment2 = New Segment2D
    With segment2
        Set .StartPoint = pointB
        Set .EndPoint = pointC
    End With
    
    Dim seg1Bisector As Line2D
    Set seg1Bisector = segment1.Bisector
    
    Dim seg2Bisector As Line2D
    Set seg2Bisector = segment2.Bisector
    
    Dim center As Point2D
    Set center = seg1Bisector.IntersectionWith(seg2Bisector)
    
    Dim Radius As Double
    Radius = center.DistanceTo(pointA)
    
    Dim newCircle As Circle2D
    Set newCircle = New Circle2D
    With newCircle
        Set .center = center
        .Radius = Radius
    End With
    
    Set MakeCircleFromPoints = newCircle
End Function
