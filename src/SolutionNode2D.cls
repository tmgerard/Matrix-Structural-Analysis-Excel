VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolutionNode2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Solution")
Option Explicit

Private Type TSolutionNode2D
    originalNode As Node2D
    displacement As Vector2D
End Type

Private this As TSolutionNode2D

Private Sub Class_Initialize()
    With this
        Set .originalNode = New Node2D
        Set .displacement = New Vector2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .originalNode = Nothing
        Set .displacement = Nothing
    End With
End Sub

Public Property Get ID() As Long
    ID = this.originalNode.ID
End Property

Public Property Get OriginalPosition() As Point2D
    Set OriginalPosition = this.originalNode.Position
End Property

Public Property Get IsConstrained() As Boolean
    IsConstrained = this.originalNode.xConstrained Or this.originalNode.yConstrained
End Property

Public Property Get Loads() As Collection
    Set Loads = this.originalNode.Loads
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = this.originalNode.LoadCount > 0
End Property

Public Property Get NetLoad() As Vector2D
    Set NetLoad = this.originalNode.NetLoad
End Property

Public Property Get displacedPosition() As Point2D
    Set displacedPosition = Me.OriginalPosition.Displaced(this.displacement)
End Property

Public Function DisplacedPositionScaled(Optional ByRef factor As Double = 1) As Point2D
    Set DisplacedPositionScaled = Me.OriginalPosition.Displaced(this.displacement, factor)
End Function

Public Sub SetSolutionNode(ByRef originalNode As Node2D, ByRef displacedPosition As Vector2D)
    Set this.originalNode = originalNode
    Set this.displacement = displacedPosition
End Sub
