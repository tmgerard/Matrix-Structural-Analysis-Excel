VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2DCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TNode2DCollection
    Nodes As Collection
End Type

Private this As TNode2DCollection

Private Sub Class_Initialize()
    With this
        Set .Nodes = New Collection
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Nodes = Nothing
    End With
End Sub

Public Function Add(ByRef Node As Node2D) As Node2DCollection
    If this.Nodes.Count = 0 Then
        Node.ID = 0
        this.Nodes.Add Node
    Else
        Dim lastNode As Node2D
        Set lastNode = this.Nodes(this.Nodes.Count)
        Node.ID = lastNode.ID + 1
        this.Nodes.Add Node
    End If
    
    Set Add = this.Nodes
End Function

Public Function Count() As Long
    Count = this.Nodes.Count
End Function

Public Function Item(ByRef nodeIndex As Long) As Node2D
    Set Item = this.Nodes.Item(nodeIndex)
End Function

Public Function Remove(ByRef nodeIndex As Long) As Node2DCollection
    Dim toRemove As Node2D
    Set toRemove = this.Nodes(nodeIndex)
    
    Dim removeID As Long
    removeID = toRemove.ID
    
    this.Nodes.Remove nodeIndex
    
    Dim Node As Node2D
    For Each Node In this.Nodes
        If Node.ID > removeID Then Node.ID = Node.ID - 1
    Next Node
    
    Set Remove = this.Nodes
End Function
