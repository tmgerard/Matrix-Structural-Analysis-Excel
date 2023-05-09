VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2DLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TNode2DLoad
    Node As Node2D
    Loads(0 To 2) As Double
End Type

Private this As TNode2DLoad

Public Property Get Node() As Node2D
    Set Node = this.Node
End Property

Public Property Set Node(ByRef value As Node2D)
    Set Node = value
End Property

Public Property Get Load(ByRef index As Long) As Double
    Load = this.Loads(index)
End Property

Public Property Let Load(ByRef index As Long, ByRef value As Double)
    this.Loads(index) = value
End Property

Public Sub SetLoad(ByRef xForce As Double, ByRef yForce As Double, ByRef zMoment As Double)
    With this
        .Loads(0) = xForce
        .Loads(1) = yForce
        .Loads(2) = zMoment
    End With
End Sub
