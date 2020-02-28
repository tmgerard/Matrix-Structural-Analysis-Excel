VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Geometry.Node")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements INode

Private Type TNode2D
    nodeName As String
    xCoordinate As Double
    yCoordinate As Double
End Type

Private this As TNode2D
Private Const CLASS_NAME As String = "Node2D"

Public Property Get Name() As String
    Name = this.nodeName
End Property

Public Property Let Name(ByVal value As String)
    this.nodeName = value
End Property

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
