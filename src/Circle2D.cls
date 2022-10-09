VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Circle2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TCircle2D
    center As Point2D
    radius As Double
End Type

Private this As TCircle2D

Private Sub Class_Initialize()
    Set this.center = New Point2D
End Sub

Private Sub Class_Terminate()
    Set this.center = Nothing
End Sub

Public Property Get center() As Point2D
    Set center = this.center
End Property

Public Property Set center(ByRef point As Point2D)
    Set this.center = point
End Property

Public Property Get radius() As Double
    radius = this.radius
End Property

Public Property Let radius(ByRef circleRadius As Double)
    ' TODO: Raise error for negative radius
    this.radius = circleRadius
End Property

'@Description "Calculates the area of the circle."
Public Function Area() As Double
Attribute Area.VB_Description = "Calculates the area of the circle."
    Area = WorksheetFunction.Pi * this.radius ^ 2
End Function

'@Description "Calculates the circumference the circle."
Public Function Circumference() As Double
Attribute Circumference.VB_Description = "Calculates the circumference the circle."
    Circumference = 2 * WorksheetFunction.Pi * this.radius
End Function

'@Description "Returns True if a given Point2D objects lies within the Circle2D object."
Public Function ContainsPoint(ByRef point As Point2D) As Boolean
Attribute ContainsPoint.VB_Description = "Returns True if a given Point2D objects lies within the Circle2D object."
    ContainsPoint = point.DistanceTo(this.center) < this.radius
End Function

'@Description "Returns a Polygon2D approximation of a Circle2D with the given number of divisions."
Public Function ToPolygon(ByRef divisions As Long) As Polygon2D
Attribute ToPolygon.VB_Description = "Returns a Polygon2D approximation of a Circle2D with the given number of divisions."
    Dim angleDelta As Double
    angleDelta = 2 * WorksheetFunction.Pi / divisions
    
    Dim pointsOnCircle As Collection
    Set pointsOnCircle = New Collection
    
    Dim division As Long
    For division = 1 To divisions
        pointsOnCircle.Add PointAtAngle(angleDelta * division)
    Next division
    
    Dim polygon As Polygon2D
    Set polygon = New Polygon2D
    Set polygon.Vertices = pointsOnCircle
    
    Set ToPolygon = polygon
End Function

Private Function PointAtAngle(ByRef angle As Double) As Point2D
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = this.center.x + this.radius * Math.cos(angle)
        .y = this.center.y + this.radius * Math.sin(angle)
    End With
    
    Set PointAtAngle = point
End Function

Public Function Equals(ByRef other As Circle2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Circle2D Then ' other is not a Circle2D
        Equals = False
    Else    ' check that x and y values are the same
        Equals = this.center.Equals(other.center) And Doubles.Equal(this.radius, other.radius)
    End If

End Function
