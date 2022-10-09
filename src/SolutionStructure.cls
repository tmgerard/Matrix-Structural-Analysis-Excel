VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolutionStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Solution")
Option Explicit

Private Type TSolutionStructure
    Bars As Collection
    Nodes As Collection
End Type

Private this As TSolutionStructure

Private Sub Class_Initialize()
    With this
        Set .Bars = New Collection
        Set .Nodes = New Collection
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Bars = Nothing
        Set .Nodes = Nothing
    End With
End Sub

Public Property Get Bars() As Collection
    Set Bars = this.Bars
End Property

Public Property Set Bars(ByRef barCollection As Collection)
    Set this.Bars = barCollection
End Property

Public Property Get Nodes() As Collection
    Set Nodes = this.Nodes
End Property

Public Property Set Nodes(ByRef nodeCollection As Collection)
    Set this.Nodes = nodeCollection
End Property

Public Function BoundingRectangle(ByRef margin As Double, Optional ByRef scalingFactor As Double = 1) As Rectangle2D

    Dim displacedPoints As Collection
    Set displacedPoints = New Collection
    
    Dim index As Long
    For index = 1 To this.Nodes.count
        displacedPoints.Add this.Nodes.Item(index).DisplacedPositionScaled(scalingFactor)
    Next index
    
    Set BoundingRectangle = MakeRectangleContainingWithMargins(displacedPoints, margin)
    
End Function

Public Function NodeReaction(ByRef Node As SolutionNode2D) As Vector2D
' TODO: Add NodeReaction Test
    Dim reaction As Vector2D
    Set reaction = New Vector2D
    reaction.u = 0
    reaction.v = 0
    
    If Node.IsConstrained Then
        Dim nodeForce As Vector2D
        Dim currentBar As SolutionBarElement2D
        Dim bar As Long
        For bar = 1 To this.Bars.count
            Set currentBar = this.Bars.Item(bar)
            If currentBar.HasNode(Node) Then
                Set nodeForce = currentBar.ForceInNode(Node)
                Set reaction = reaction.Add(nodeForce)
            End If
        Next bar
    Else
        With reaction
            .u = 0
            .v = 0
        End With
    End If
    
    If Node.IsLoaded Then
        Set reaction = reaction.Add(Node.NetLoad.Opposite)
    End If
    
    
    Set NodeReaction = reaction

End Function
