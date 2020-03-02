VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coordinate3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.Geometry.Coordinates")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements ICoordinate

Private Type TCoordinate3D
    xCoordinate As Double
    yCoordinate As Double
    zCoordinate As Double
End Type

Private this As TCoordinate3D
Private Const CLASS_NAME As String = "Coordinate3D"

Public Property Get X() As Double
    X = this.xCoordinate
End Property

Public Property Let X(ByVal value As Double)
    this.xCoordinate = value
End Property

Public Property Get Y() As Double
    Y = this.yCoordinate
End Property

Public Property Let Y(ByVal value As Double)
    this.yCoordinate = value
End Property

Public Property Get Z() As Double
    Z = this.zCoordinate
End Property

Public Property Let Z(ByVal value As Double)
    this.zCoordinate = value
End Property

Public Sub SetCoordinates(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    this.xCoordinate = X
    this.yCoordinate = Y
    this.zCoordinate = Z
End Sub

Public Function Distance(ByVal other As ICoordinate) As Double
    
    Dim xDifference As Double
    xDifference = other.X - this.xCoordinate
    
    Dim yDifference As Double
    yDifference = other.Y - this.yCoordinate
    
    Dim zDifference As Double
    zDifference = other.Z - this.zCoordinate
    
    Distance = Math.Sqr(xDifference ^ 2 + yDifference ^ 2 + zDifference ^ 2)
    
End Function

'ICoordinate Implementation

Private Function ICoordinate_Distance(ByVal other As ICoordinate) As Double
    ICoordinate_Distance = Me.Distance(other)
End Function

Private Property Get ICoordinate_X() As Double
    ICoordinate_X = Me.X
End Property

Private Property Get ICoordinate_Y() As Double
    ICoordinate_Y = Me.Y
End Property

Private Property Get ICoordinate_Z() As Double
    ICoordinate_Z = Me.Z
End Property
