VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AffineTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TAffineTransform
    ScaleX As Double
    ScaleY As Double
    translateX As Double
    translateY As Double
    shearX As Double
    shearY As Double
End Type

Private this As TAffineTransform

Private Sub Class_Initialize()
    With this
        .ScaleX = 1
        .ScaleY = 1
        .translateX = 0
        .translateY = 0
        .shearX = 0
        .shearY = 0
    End With
End Sub

Public Property Get ScaleX() As Double
    ScaleX = this.ScaleX
End Property

Public Property Let ScaleX(ByRef value As Double)
    this.ScaleX = value
End Property

Public Property Get ScaleY() As Double
    ScaleY = this.ScaleY
End Property

Public Property Let ScaleY(ByRef value As Double)
    this.ScaleY = value
End Property

Public Property Get translateX() As Double
    translateX = this.translateX
End Property

Public Property Let translateX(ByRef value As Double)
    this.translateX = value
End Property

Public Property Get translateY() As Double
    translateY = this.translateY
End Property

Public Property Let translateY(ByRef value As Double)
    this.translateY = value
End Property

Public Property Get shearX() As Double
    shearX = this.shearX
End Property

Public Property Let shearX(ByRef value As Double)
    this.shearX = value
End Property

Public Property Get shearY() As Double
    shearY = this.shearY
End Property

Public Property Let shearY(ByRef value As Double)
    this.shearY = value
End Property

Public Sub SetTransform(ByRef ScaleX As Double, ByRef ScaleY As Double, _
                        ByRef translateX As Double, ByRef translateY As Double, _
                        ByRef shearX As Double, ByRef shearY As Double)

    With this
        .ScaleX = ScaleX
        .ScaleY = ScaleY
        .translateX = translateX
        .translateY = translateY
        .shearX = shearX
        .shearY = shearY
    End With

End Sub

Public Function ApplyToPoint(ByRef point As Point2D) As Point2D

    Dim transformedPoint As Point2D
    Set transformedPoint = New Point2D
    With transformedPoint
        .x = (this.ScaleX * point.x) + (this.shearX * point.y) + this.translateX
        .y = (this.shearY * point.x) + (this.ScaleY * point.y) + this.translateY
    End With
    
    Set ApplyToPoint = transformedPoint
End Function

Public Function ApplyToSegment(ByRef segment As Segment2D) As Segment2D

    Dim transformedSegment As Segment2D
    Set transformedSegment = New Segment2D
    With transformedSegment
        Set .StartPoint = Me.ApplyToPoint(segment.StartPoint)
        Set .EndPoint = Me.ApplyToPoint(segment.EndPoint)
    End With
    
    Set ApplyToSegment = transformedSegment

End Function

Public Function ApplyToPolygon(ByRef polygon As Polygon2D) As Polygon2D
    
    Dim newVertices As Collection
    Set newVertices = New Collection

    Dim index As Long
    For index = 1 To polygon.Vertices.Count
        newVertices.Add Me.ApplyToPoint(polygon.Vertices.Item(index))
    Next index
    
    Dim transformedPolygon As Polygon2D
    Set transformedPolygon = New Polygon2D
    With transformedPolygon
        Set .Vertices = newVertices
    End With
    
    Set ApplyToPolygon = transformedPolygon
    
End Function

Public Function ApplyToRectangle(ByRef rectangle As Rectangle2D) As Polygon2D
    Set ApplyToRectangle = Me.ApplyToPolygon(rectangle.ToPolygon)
End Function

Public Function ApplyToCircle(ByRef circ As Circle2D, Optional ByRef divisions As Long = 30) As Polygon2D
    Set ApplyToCircle = Me.ApplyToPolygon(circ.ToPolygon(divisions))
End Function

Public Function Concatenate(ByRef other As AffineTransform) As AffineTransform
    
    Dim newTransform As AffineTransform
    Set newTransform = New AffineTransform
    With newTransform
        .ScaleX = other.ScaleX * this.ScaleX + other.shearX * this.shearY
        .ScaleY = other.shearY * this.shearX + other.ScaleY * this.ScaleY
        .translateX = other.ScaleX * this.translateX + other.shearX * this.translateY + other.translateX
        .translateY = other.shearY * this.translateX + other.ScaleY * this.translateY + other.translateY
        .shearX = other.ScaleX * this.shearX + other.shearX * this.ScaleY
        .shearY = other.shearY * this.ScaleX + other.ScaleY * this.shearY
    End With
    
    Set Concatenate = newTransform
    
End Function

Public Function Inverse() As AffineTransform

    Dim denom As Double
    denom = this.ScaleX * this.ScaleY - this.shearX * this.shearY
    
    Dim inverseTrans As AffineTransform
    Set inverseTrans = New AffineTransform
    With inverseTrans
        .ScaleX = this.ScaleY / denom
        .ScaleY = this.ScaleX / denom
        .translateX = (this.translateY * this.shearX - this.ScaleY * this.translateX) / denom
        .translateY = (this.translateX * this.shearY - this.ScaleX * this.translateY) / denom
        .shearX = -this.shearX / denom
        .shearY = -this.shearY / denom
    End With
    
    Set Inverse = inverseTrans

End Function

Public Function Equals(ByRef other As AffineTransform) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is AffineTransform Then ' other is not a AffineTransform
        Equals = False
    Else
        Equals = Doubles.Equal(this.ScaleX, other.ScaleX) And _
                 Doubles.Equal(this.ScaleY, other.ScaleY) And _
                 Doubles.Equal(this.translateX, other.translateX) And _
                 Doubles.Equal(this.translateY, other.translateY) And _
                 Doubles.Equal(this.shearX, other.shearX) And _
                 Doubles.Equal(this.shearY, other.shearY)
    End If

End Function

