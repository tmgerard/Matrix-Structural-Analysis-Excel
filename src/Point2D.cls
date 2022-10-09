VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TPoint2D
    x As Double
    y As Double
End Type

Private this As TPoint2D

Public Property Get x() As Double
    x = this.x
End Property

Public Property Let x(ByRef value As Double)
    this.x = value
End Property

Public Property Get y() As Double
    y = this.y
End Property

Public Property Let y(ByRef value As Double)
    this.y = value
End Property

'@Description "Calculates distance between two Point2D objects."
Public Function DistanceTo(ByRef other As Point2D)
Attribute DistanceTo.VB_Description = "Calculates distance between two Point2D objects."
    Dim delta_x As Double
    delta_x = other.x - this.x
    
    Dim delta_y As Double
    delta_y = other.y - this.y
    
    DistanceTo = Math.Sqr(delta_x ^ 2 + delta_y ^ 2)
End Function

'@Description "Adds x and y coordinate values to given Point2D object and returns new Point2D object."
Public Function Add(ByRef other As Point2D) As Point2D
Attribute Add.VB_Description = "Adds x and y coordinate values to given Point2D object and returns new Point2D object."
    Dim addPoint As Point2D
    Set addPoint = New Point2D
    With addPoint
        .x = this.x + other.x
        .y = this.y + other.y
    End With
    
    Set Add = addPoint
End Function

'@Description "Subtracts x and y coordinate values to given Point2D object and returns new Vector2D object."
Public Function Subtract(ByRef other As Point2D) As Vector2D
Attribute Subtract.VB_Description = "Subtracts x and y coordinate values to given Point2D object and returns new Point2D object."
    Dim subtractVector As Vector2D
    Set subtractVector = New Vector2D
    With subtractVector
        .u = this.x - other.x
        .v = this.y - other.y
    End With
    
    Set Subtract = subtractVector
End Function

'@Description "Returns a new Point2D object that is the original object displaced by a Vector2D."
Public Function Displaced(ByRef Vector As Vector2D, Optional ByRef times As Double = 1) As Point2D
Attribute Displaced.VB_Description = "Returns a new Point2D object that is the original object displaced by a Vector2D."
    Dim scaledVector As Vector2D
    Set scaledVector = Vector.ScaledBy(times)
    
    Dim displacedPoint As Point2D
    Set displacedPoint = New Point2D
    With displacedPoint
        .x = this.x + scaledVector.u
        .y = this.y + scaledVector.v
    End With
    
    Set Displaced = displacedPoint
End Function

Public Function Equals(ByRef other As Point2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Point2D Then ' other is not a Point2D
        Equals = False
    Else    ' check that x and y values are the same
        Equals = Doubles.Equal(this.x, other.x) And Doubles.Equal(this.y, other.y)
    End If

End Function
