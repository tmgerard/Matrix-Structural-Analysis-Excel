VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Segment2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TSegement
    Start As Point2D
    End As Point2D
End Type

Private this As TSegement

Private Sub Class_Initialize()
    With this
        Set .Start = New Point2D
        Set .End = New Point2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Start = Nothing
        Set .End = Nothing
    End With
End Sub

Public Property Get StartPoint() As Point2D
    Set StartPoint = this.Start
End Property

Public Property Set StartPoint(ByRef point As Point2D)
    Set this.Start = point
End Property

Public Property Get EndPoint() As Point2D
    Set EndPoint = this.End
End Property

Public Property Set EndPoint(ByRef point As Point2D)
    Set this.End = point
End Property

'@Description "Return Vector2D object between Segment end points."
Public Function DirectionVector() As Vector2D
Attribute DirectionVector.VB_Description = "Return Vector2D object between Segment end points."
    Set DirectionVector = Vector2DFactory.MakeVectorBetween(this.Start, this.End)
End Function

'@Description "Return Vector2D object with unit length in direction of Segment."
Public Function UnitDirectionVector() As Vector2D
Attribute UnitDirectionVector.VB_Description = "Return Vector2D object with unit length in direction of Segment."
    Set UnitDirectionVector = Vector2DFactory.MakeUnitVectorBetween(this.Start, this.End)
End Function

'@Description "Return Vector2D object in direction perpendicular to Segment."
Public Function NormalUnitVector() As Vector2D
Attribute NormalUnitVector.VB_Description = "Return Vector2D object in direction perpendicular to Segment."
    Set NormalUnitVector = DirectionVector.Perpendicular
End Function

'@Description "Returns the length of the Segment object."
Public Function Length() As Double
Attribute Length.VB_Description = "Returns the length of the Segment object."
    Length = this.Start.DistanceTo(this.End)
End Function

'@Description "Returns Point2D object corresponding to a location on the line segment given as a ratio of the segment length."
Public Function PointAt(ByRef ratio As Double) As Point2D
Attribute PointAt.VB_Description = "Returns Point2D object corresponding to a location on the line segment given as a ratio of the segment length."
    Segment2DRatio.EnsureValidRatio ratio:=ratio
    Set PointAt = this.Start.Displaced(Me.DirectionVector, ratio)
End Function

'@Description "Returns Point2D object at the center of the segment."
Public Function PointAtMiddle() As Point2D
Attribute PointAtMiddle.VB_Description = "Returns Point2D object at the center of the segment."
    Set PointAtMiddle = PointAt(Segment2DRatio.MID)
End Function

'@Description "Returns Point2D object corresponding to the closest point on the segment to a point outside the segment."
Public Function ClosestPointTo(ByRef outsidePoint As Point2D) As Point2D
Attribute ClosestPointTo.VB_Description = "Returns Point2D object corresponding to the closest point on the segment to a point outside the segment."
    Dim vectorToPoint As Vector2D
    Set vectorToPoint = Vector2DFactory.MakeVectorBetween(this.Start, outsidePoint)
    
    Dim unitVectorAlongSegment As Vector2D
    Set unitVectorAlongSegment = Me.UnitDirectionVector
    
    Dim projectionOnSegment As Double
    projectionOnSegment = vectorToPoint.ProjectionOver(unitVectorAlongSegment)
    
    If projectionOnSegment < 0 Then
        Set ClosestPointTo = this.Start
    ElseIf projectionOnSegment > Me.Length Then
        Set ClosestPointTo = this.End
    Else
        Set ClosestPointTo = this.Start.Displaced(unitVectorAlongSegment, projectionOnSegment)
    End If
    
End Function

'@Description "Returns the distance between a point in space to its closest point on the segment."
Public Function DistanceTo(ByRef outsidePoint As Point2D)
Attribute DistanceTo.VB_Description = "Returns the distance between a point in space to its closes point on the segment."
    Set DistanceTo = outsidePoint.DistanceTo(Me.ClosestPointTo(outsidePoint))
End Function

'@Description "Returns Point2D object at intersection between two Segment2D objects."
Public Function IntersectionWith(ByRef other As Segment2D) As Point2D
Attribute IntersectionWith.VB_Description = "Returns Point2D object at intersection between two Segment2D objects."
    Dim d1 As Vector2D
    Set d1 = Me.DirectionVector
    
    Dim d2 As Vector2D
    Set d2 = other.DirectionVector
    
    If d1.IsParallelTo(d2) Then
        Set IntersectionWith = Nothing
        Exit Function
    End If
    
    Dim Cross As Double
    Cross = d1.Cross(d2)
    
    Dim delta As Vector2D
    Set delta = other.StartPoint.Subtract(this.Start)
    
    Dim t1 As Double
    t1 = (delta.u * d2.v - delta.v * d2.u) / Cross
    
    Dim t2 As Double
    t2 = (delta.u * d1.v - delta.v * d1.u) / Cross
    
    If Segment2DRatio.IsValid(t1) And Segment2DRatio.IsValid(t2) Then
        Set IntersectionWith = Me.PointAt(t1)
    Else
        Set IntersectionWith = Nothing
    End If
    
End Function

Public Function Equals(ByRef other As Segment2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Segment2D Then ' other is not a Vector2D
        Equals = False
    Else    ' check that segment end points are the same
        Equals = this.Start.Equals(other.StartPoint) And this.End.Equals(other.EndPoint)
    End If

End Function

'@Description "Returns Line2D object perpendicular to segment through the point at the middle of the segment."
Public Function Bisector() As Line2D
Attribute Bisector.VB_Description = "Returns Line2D object perpendicular to segment through the point at the middle of the segment."
    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        Set .Base = Me.PointAtMiddle
        Set .direction = Me.NormalUnitVector
    End With
    
    Set Bisector = newLine
End Function
