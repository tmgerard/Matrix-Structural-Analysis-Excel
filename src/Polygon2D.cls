VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Polygon2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TPolygon2D
    points As Collection
End Type

Private this As TPolygon2D
Private Const MinVertices As Integer = 3

Public Enum Polygon2DErrors
    NotEnoughVertices = 200 + vbObjectError
End Enum

Private Sub Class_Initialize()
    Set this.points = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.points = Nothing
End Sub

Public Property Get Vertices() As Collection
    Set Vertices = this.points
End Property

Public Property Set Vertices(ByRef points As Collection)

    If points.count < 3 Then
        Err.Raise Number:=Polygon2DErrors.NotEnoughVertices, _
                  Source:="Polygon2D.Vertices (Set)", _
                  Description:=MinVertices & " or more vertices required for Polygon2D object"
    End If

    Dim point As Variant
    For Each point In points
        If TypeOf point Is Point2D Then
            this.points.Add point
        Else    ' Throw vba type mismatch error code
            Err.Raise Number:=13, _
                      Source:="Polygon2D.Vertices (Set)", _
                      Description:="Polygon2D vertex type Mismatch. Point2D objects required."

        End If
    Next point

End Property

'@Description "Returns a collection of Segment2D objects representing the sides of the Polygon2D object."
Public Function Sides() As Collection
Attribute Sides.VB_Description = "Returns a collection of Segment2D objects representing the sides of the Polygon2D object."

    Dim segments As Collection
    Set segments = New Collection
    
    Dim segment As Segment2D
    Dim point As Long
    For point = 1 To this.points.count
    
        Dim nextPoint As Long
        If point < this.points.count Then
            nextPoint = point + 1
        Else ' Go back to first point index
            nextPoint = 1
        End If
        
        Set segment = New Segment2D
        With segment
            Set .StartPoint = this.points(point)
            Set .EndPoint = this.points(nextPoint)
        End With
        
        segments.Add segment
    Next point
    
    Set Sides = segments

End Function

'@Description "Point2D object representing the centroid of the Polygon2D object."
Public Function Centroid() As Point2D
Attribute Centroid.VB_Description = "Point2D object representing the centroid of the Polygon2D object."
    Dim sum_x As Double
    sum_x = 0
    
    Dim sum_y As Double
    sum_y = 0
    
    Dim point As Point2D
    For Each point In this.points
        sum_x = sum_x + point.x
        sum_y = sum_y + point.y
    Next point
    
    Dim center As Point2D
    Set center = New Point2D
    With center
        .x = sum_x / this.points.count
        .y = sum_y / this.points.count
    End With
    
    Set Centroid = center
End Function

'@Description "Returns True if the given Point2D object lies within the Polygon2D object."
Public Function ContainsPoint(ByRef point As Point2D) As Boolean
Attribute ContainsPoint.VB_Description = "Returns True if the given Point2D object lies within the Polygon2D object."
' Winding algorithm. Sum of angle from point to vectors equal 2*pi
' when the point lies inside the polygon and zero if outside.

    Dim vertexIndex As Long
    For vertexIndex = 1 To this.points.count
        If point.Equals(this.points.Item(vertexIndex)) Then
            ContainsPoint = True
            Exit Function
        End If
    Next vertexIndex

    Dim pointToVertexVectors As Collection
    Set pointToVertexVectors = New Collection
    
    Dim vertex As Point2D
    For Each vertex In this.points
        pointToVertexVectors.Add MakeVectorBetween(point, vertex)
    Next vertex
    
    Dim angleSum As Double
    angleSum = 0
    
    Dim v1 As Vector2D
    Dim v2 As Vector2D
    Dim vectorIndex As Long
    For vectorIndex = 1 To pointToVertexVectors.count
    
        Dim nextVector As Long
        If vectorIndex < pointToVertexVectors.count Then
            nextVector = vectorIndex + 1
        Else
            nextVector = 1
        End If
        
        Set v1 = pointToVertexVectors.Item(vectorIndex)
        Set v2 = pointToVertexVectors.Item(nextVector)
        
        angleSum = angleSum + v1.AngleTo(v2)
    
    Next vectorIndex
    
    ContainsPoint = Doubles.Equal(angleSum, 2 * WorksheetFunction.Pi)
    
End Function

Public Function Equals(ByRef other As Polygon2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Polygon2D Then ' other is not a Vector2D
        Equals = False
    Else    ' check that polygon vertices are equal
        If Not Me.Vertices.count = other.Vertices.count Then
            Equals = False
        Else
            Dim vertexIndex As Long
            For vertexIndex = 1 To this.points.count
                If Not this.points.Item(vertexIndex).Equals(other.Vertices.Item(vertexIndex)) Then
                    Equals = False
                    Exit Function
                End If
            Next vertexIndex
            
            Equals = True
        End If
    End If

End Function
