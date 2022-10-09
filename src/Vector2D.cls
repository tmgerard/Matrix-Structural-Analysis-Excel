VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TVector2D
    u As Double
    v As Double
End Type

Private this As TVector2D

Public Property Get u() As Double
    u = this.u
End Property

Public Property Let u(ByRef value As Double)
    this.u = value
End Property

Public Property Get v() As Double
    v = this.v
End Property

Public Property Let v(ByRef value As Double)
    this.v = value
End Property

'@Description "Adds a Vector2D to object and returns a new Vector2D object."
Public Function Add(ByRef other As Vector2D) As Vector2D
Attribute Add.VB_Description = "Adds a Vector2D to object and returns a new Vector2D object."
    Dim addVector As Vector2D
    Set addVector = New Vector2D
    With addVector
        .u = this.u + other.u
        .v = this.v + other.v
    End With
    
    Set Add = addVector
End Function

'@Description "Subtracts a Vector2D from object and returns new Vector2D object."
Public Function Subtract(ByRef other As Vector2D) As Vector2D
Attribute Subtract.VB_Description = "Subtracts a Vector2D from object and returns new Vector2D object."
    Dim subtractVector As Vector2D
    Set subtractVector = New Vector2D
    With subtractVector
        .u = this.u - other.u
        .v = this.v - other.v
    End With
    
    Set Subtract = subtractVector
End Function

'@Description "Scales Vector2D object by a value and returns new Vector2D object."
Public Function ScaledBy(ByRef factor As Double) As Vector2D
Attribute ScaledBy.VB_Description = "Scales Vector2D object by a value and returns new Vector2D object."
    Dim scaledVector As Vector2D
    Set scaledVector = New Vector2D
    With scaledVector
        .u = factor * this.u
        .v = factor * this.v
    End With
    
    Set ScaledBy = scaledVector
End Function

'@Description "Returns length of Vector2D object."
Public Function Norm() As Double
Attribute Norm.VB_Description = "Returns length of Vector2D object."
    Norm = Math.Sqr(this.u ^ 2 + this.v ^ 2)
End Function

'@Description "Returns True if vector length is normalized (length = 1)."
Public Function IsNormal() As Boolean
Attribute IsNormal.VB_Description = "Returns True if vector length is normalized (length = 1)."
    IsNormal = Doubles.IsEffectivelyOne(Me.Norm)
End Function

'@Description "Normalizes Vector2D object and returns new Vector2D object."
Public Function Normalized() As Vector2D
Attribute Normalized.VB_Description = "Normalizes Vector2D object and returns new Vector2D object."
    Set Normalized = Me.ScaledBy(1# / Me.Norm)
End Function

'@Description "Returns new Vector2D object with a specified length in the direction of the original Vector2D object."
Public Function WithLength(ByRef Length As Double) As Vector2D
Attribute WithLength.VB_Description = "Returns new Vector2D object with a specified length in the direction of the original Vector2D object."
    Set WithLength = Me.Normalized.ScaledBy(Length)
End Function

'@Description "Returns the dot product of to Vector2D objects."
Public Function Dot(ByRef other As Vector2D) As Double
Attribute Dot.VB_Description = "Returns the dot product of to Vector2D objects."
    Dot = (this.u * other.u) + (this.v * other.v)
End Function

'@Description "Returns the projection of a Vector2D object over another Vector2D object."
Public Function ProjectionOver(ByRef direction As Vector2D) As Double
Attribute ProjectionOver.VB_Description = "Returns the projection of a Vector2D object over another Vector2D object."
    ProjectionOver = Me.Dot(direction.Normalized)
End Function

'@Description "Returns the cross product of two Vector2D objects."
Public Function Cross(ByRef other As Vector2D) As Double
Attribute Cross.VB_Description = "Returns the cross product of two Vector2D objects."
    ' The cross product for 2D vectors is used to determine
    ' the rotational direction of angles. A positive value
    ' will indicate a counterclockwise rotation while a negative
    ' value will indicate a clockwise rotation. A value of zero
    ' indicates that vectors are parallel.
    Cross = (this.u * other.v) - (this.v * other.u)
End Function

'@Description "Returns True if Vector2D objects are parallel."
Public Function IsParallelTo(ByRef other As Vector2D) As Boolean
Attribute IsParallelTo.VB_Description = "Returns True if Vector2D objects are parallel."
    IsParallelTo = Doubles.IsEffectivelyZero(Me.Cross(other))
End Function

'@Description "Returns True if Vector2D objects are perpendicular."
Public Function IsPerpendicularTo(ByRef other As Vector2D) As Boolean
Attribute IsPerpendicularTo.VB_Description = "Returns True if Vector2D objects are perpendicular."
    IsPerpendicularTo = Doubles.IsEffectivelyZero(Me.Dot(other))
End Function

'@Description "Returns the magnitude of the angle between two Vector2D objects."
Public Function AngleValueTo(ByRef other As Vector2D) As Double
Attribute AngleValueTo.VB_Description = "Returns the magnitude of the angle between two Vector2D objects."
    Dim dotProduct As Double
    dotProduct = Me.Dot(other)
    
    Dim normProduct As Double
    normProduct = Me.Norm * other.Norm
    
    AngleValueTo = WorksheetFunction.Acos(dotProduct / normProduct)
End Function

'@Description "Returns the angle between two Vector2D objects. A positive value indicates a counterclockwise rotation."
Public Function AngleTo(ByRef other As Vector2D) As Double
Attribute AngleTo.VB_Description = "Returns the angle between two Vector2D objects. A positive value indicates a counterclockwise rotation."
    Dim value As Double
    value = Me.AngleValueTo(other)
    
    Dim crossProduct As Double
    crossProduct = Me.Cross(other)
    
    AngleTo = Math.Sgn(crossProduct) * value
End Function

'@Description "Returns new Vector2D object that is a rotation of the original vector by a given angle in radians."
Public Function Rotated(ByRef radians As Double) As Vector2D
Attribute Rotated.VB_Description = "Returns new Vector2D object that is a rotation of the original vector by a given angle in radians."
    Dim cos As Double
    cos = Math.cos(radians)
    
    Dim sin As Double
    sin = Math.sin(radians)
    
    Dim rotatedVector As Vector2D
    Set rotatedVector = New Vector2D
    With rotatedVector
        .u = this.u * cos - this.v * sin
        .v = this.u * sin + this.v * cos
    End With
    
    Set Rotated = rotatedVector
End Function

'@Description "Returns Vector2D object perpendicular to the original."
Public Function Perpendicular() As Vector2D
Attribute Perpendicular.VB_Description = "Returns Vector2D object perpendicular to the original."
    Dim perpendicularVector As Vector2D
    Set perpendicularVector = New Vector2D
    With perpendicularVector
        .u = -this.v
        .v = this.u
    End With
    
    Set Perpendicular = perpendicularVector
End Function

'@Description "Returns Vector2D object with equal magnitude and in the opposite direction of the original."
Public Function Opposite() As Vector2D
Attribute Opposite.VB_Description = "Returns Vector2D object with equal magnitude and in the opposite direction of the original."
    Dim oppositeVector As Vector2D
    Set oppositeVector = New Vector2D
    With oppositeVector
        .u = -this.u
        .v = -this.v
    End With
    
    Set Opposite = oppositeVector
End Function

'@Description "Returns the Vector2D's direction sine value."
Public Function Sine() As Double
Attribute Sine.VB_Description = "Returns the Vector2D's direction sine value."
    Sine = this.v / Me.Norm
End Function

'@Description "Returns the Vector2D's direction cosine value."
Public Function Cosine() As Double
Attribute Cosine.VB_Description = "Returns the Vector2D's direction cosine value."
    Cosine = this.u / Me.Norm
End Function

Public Function Equals(ByRef other As Vector2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Vector2D Then ' other is not a Vector2D
        Equals = False
    Else    ' check that u and v values are the same
        Equals = Doubles.Equal(this.u, other.u) And Doubles.Equal(this.v, other.v)
    End If

End Function
