VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rectangle2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TRectangle2D
    Origin As Point2D
    size As Rectangle2DSize
End Type

Private this As TRectangle2D

Private Sub Class_Initialize()
    With this
        Set .Origin = New Point2D
        Set .size = New Rectangle2DSize
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Origin = Nothing
        Set .size = Nothing
    End With
End Sub

Public Property Get Origin() As Point2D
    Set Origin = this.Origin
End Property

Public Property Set Origin(ByRef originPoint As Point2D)
    Set this.Origin = originPoint
End Property

Public Property Get size() As Rectangle2DSize
    Set size = this.size
End Property

Public Property Set size(ByRef rectangleSize As Rectangle2DSize)
    Set this.size = rectangleSize
End Property

Public Property Get Left() As Double
    Left = this.Origin.x
End Property

Public Property Get Right() As Double
    Right = this.Origin.x + this.size.Width
End Property

Public Property Get Bottom() As Double
    Bottom = this.Origin.y
End Property

Public Property Get Top() As Double
    Top = this.Origin.y + this.size.Height
End Property

Public Function Area() As Double
    Area = this.size.Height * this.size.Width
End Function

Public Function Perimeter() As Double
    Perimeter = 2 * (this.size.Height + this.size.Width)
End Function

Public Function ContainsPoint(ByRef point As Point2D) As Boolean
    ContainsPoint = Me.Left < point.x And point.x < Me.Right And _
                    Me.Bottom < point.y And point.y < Me.Top
End Function

Public Function ToPolygon() As Polygon2D
    Dim point1 As New Point2D
    Set point1 = New Point2D
    point1.x = Me.Right
    point1.y = Me.Bottom
    
    Dim point2 As New Point2D
    Set point2 = New Point2D
    point2.x = Me.Right
    point2.y = Me.Top
    
    Dim point3 As New Point2D
    Set point3 = New Point2D
    point3.x = Me.Left
    point3.y = Me.Top
    
    Dim points As Collection
    Set points = New Collection
    With points
        .Add Me.Origin
        .Add point1
        .Add point2
        .Add point3
    End With
    
    Dim polygon As Polygon2D
    Set polygon = New Polygon2D
    Set polygon.Vertices = points
    
    Set ToPolygon = polygon
End Function

Public Function Equals(ByRef other As Rectangle2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Rectangle2D Then ' other is not a Rectangle2D
        Equals = False
    Else    ' check that origin and sizes are equal
        Equals = this.Origin.Equals(other.Origin) And this.size.Equals(other.size)
    End If

End Function

Public Function IntersectsWith(ByRef other As Rectangle2D) As Rectangle2D

    Dim horizontalOverlap As OpenInterval
    Set horizontalOverlap = New OpenInterval
    Set horizontalOverlap = HorizontalOverlapWith(other)
    
    If horizontalOverlap Is Nothing Then
        Set IntersectsWith = Nothing
        Exit Function
    End If
    
    Dim verticalOverlap As OpenInterval
    Set verticalOverlap = New OpenInterval
    Set verticalOverlap = VerticalOverlapWith(other)
    
    If verticalOverlap Is Nothing Then
        Set IntersectsWith = Nothing
        Exit Function
    End If
    
    Dim rectPoint As Point2D
    Set rectPoint = New Point2D
    With rectPoint
        .x = horizontalOverlap.IntStart
        .y = verticalOverlap.IntStart
    End With
    
    Dim rectSize As Rectangle2DSize
    Set rectSize = New Rectangle2DSize
    With rectSize
        .Width = horizontalOverlap.Length
        .Height = verticalOverlap.Length
    End With
    
    Dim overlapRect As Rectangle2D
    Set overlapRect = New Rectangle2D
    With overlapRect
        Set .Origin = rectPoint
        Set .size = rectSize
    End With
    
    Set IntersectsWith = overlapRect
    
End Function

Private Function HorizontalOverlapWith(ByRef other As Rectangle2D) As OpenInterval
    Dim selfInterval As OpenInterval
    Set selfInterval = New OpenInterval
    selfInterval.SetInterval Me.Left, Me.Right
    
    Dim otherInterval As OpenInterval
    Set otherInterval = New OpenInterval
    otherInterval.SetInterval other.Left, other.Right
    
    Set HorizontalOverlapWith = selfInterval.ComputeOverlapWith(otherInterval)
End Function

Private Function VerticalOverlapWith(ByRef other As Rectangle2D) As OpenInterval
    Dim selfInterval As OpenInterval
    Set selfInterval = New OpenInterval
    selfInterval.SetInterval Me.Bottom, Me.Top
    
    Dim otherInterval As OpenInterval
    Set otherInterval = New OpenInterval
    otherInterval.SetInterval other.Bottom, other.Top
    
    Set VerticalOverlapWith = selfInterval.ComputeOverlapWith(otherInterval)
End Function
