VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlaneFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Const MaxDofsPerNode As Long = 3
Private Const MaxDofsPerElement As Long = 6

Private Type TPlaneFrame
    Nodes As Node2DCollection
    Elements As LinearElementCollection
    NodeLoads As Collection
    ElementLoads As Collection
End Type

Private this As TPlaneFrame

Private Sub Class_Initialize()
    With this
        .Nodes = New Node2DCollection
        .Elements = New LinearElementCollection
        .NodeLoads = New Collection
        .ElementLoads = New Collection
    End With
End Sub

Private Sub Class_Terminate()
    With this
        .Nodes = Nothing
        .Elements = Nothing
        .NodeLoads = Nothing
        .ElementLoads = Nothing
    End With
End Sub

Public Property Get Nodes() As Node2DCollection
    Set Nodes = this.Nodes
End Property

Public Property Get Elements() As LinearElementCollection
    Set Elements = this.Elements
End Property

Public Sub AddNode(ByRef strNode As Node2D)
    this.Nodes.Add strNode
End Sub

Public Sub AddElement(ByRef StartNode As Node2D, ByRef EndNode As Node2D, ByRef Area As Double, _
    ByRef ElasticMod As Double, ByRef Intertia As Double)
    this.Elements.Add MakeBeamElement2D(0, StartNode, EndNode, Area, ElasticMod, Inertia)
End Sub

Public Sub AddNodeLoad(ByRef strNode As Node2D, ByRef forceX As Double, ByRef forceY As Double, ByRef momentZ As Double)
    Dim loading As Node2DLoad
    Set loading = New Node2DLoad
    
    With loading
        Set .Node = strNode
        .SetLoad forceX, forceY, momentZ
    End With
    
    this.NodeLoads.Add loading
End Sub

Public Function DofsPerNode() As Long
    DofsPerNode = MaxDofsPerNode
End Function

Public Function DofsPerElement() As Long
    DofsPerElement = MaxDofsPerElement
End Function

Public Function NodeCount() As Long
    NodeCount = this.Nodes.Count
End Function

Public Function ElementCount() As Long
    ElementCount = this.Elements.Count
End Function
