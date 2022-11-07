VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolutionBarElement2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Solution")
Option Explicit

Private Type TStructureBarElement2D
    OriginalBar As BarElement2D
    StartNode As SolutionNode2D
    EndNode As SolutionNode2D
End Type

Private this As TStructureBarElement2D

Public Enum SolutionBarElement2DErrors
    NodeMismatch = 700 + vbObjectError
    UnkownNode
End Enum

Private Sub Class_Initialize()
    With this
        Set .OriginalBar = New BarElement2D
        Set .StartNode = New SolutionNode2D
        Set .EndNode = New SolutionNode2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .OriginalBar = Nothing
        Set .StartNode = Nothing
        Set .EndNode = Nothing
    End With
End Sub

Public Sub SetSolutionBar(ByRef OriginalBar As BarElement2D, ByRef StartNode As SolutionNode2D, ByRef EndNode As SolutionNode2D)

    If Not OriginalBar.StartNode.ID = StartNode.ID Then
        Err.Raise Number:=SolutionBarElement2DErrors.NodeMismatch, _
                  source:="SolutionBarElement2D.SetSolutionBar", _
                  Description:="Wrong start node."
    End If
    
    If Not OriginalBar.EndNode.ID = EndNode.ID Then
        Err.Raise Number:=SolutionBarElement2DErrors.NodeMismatch, _
                  source:="SolutionBarElement2D.SetSolutionBar", _
                  Description:="Wrong end node."
    End If
    
    With this
        Set .OriginalBar = OriginalBar
        Set .StartNode = StartNode
        Set .EndNode = EndNode
    End With
End Sub

Public Property Get OriginalBar() As BarElement2D
    Set OriginalBar = this.OriginalBar
End Property

Public Property Get StartNode() As SolutionNode2D
    Set StartNode = this.StartNode
End Property

Public Property Get EndNode() As SolutionNode2D
    Set EndNode = this.EndNode
End Property

Public Property Get ID() As Long
    ID = this.OriginalBar.ID
End Property

Public Property Get Area() As Double
    Area = this.OriginalBar.Area
End Property

Public Property Get ElasticModulus() As Double
    ElasticModulus = this.OriginalBar.ElasticModulus
End Property

Public Function OriginalBarGeometry() As Segment2D
    Set OriginalBarGeometry = this.OriginalBar.Geometry
End Function

Public Function FinalGeometry() As Segment2D
    Dim segment As Segment2D
    Set segment = New Segment2D
    With segment
        Set .StartPoint = this.StartNode.displacedPosition
        Set .EndPoint = this.EndNode.displacedPosition
    End With
    
    Set FinalGeometry = segment
End Function

Public Function OriginalLength() As Double
    OriginalLength = Me.OriginalBarGeometry.Length
End Function

Public Function FinalLength() As Double
    FinalLength = Me.FinalGeometry.Length
End Function

Public Function elongation() As Double
    elongation = Me.FinalLength - Me.OriginalLength
End Function

Public Function strain() As Double
    strain = Me.elongation / Me.OriginalLength
End Function

Public Function stress() As Double
    stress = Me.ElasticModulus * Me.strain
End Function

Public Function InternalForceValue() As Double
    InternalForceValue = Me.stress * Me.Area
End Function

Public Function ForceInNode(ByRef Node As SolutionNode2D) As Vector2D
    
    Dim force As Vector2D
    
    If Node Is this.StartNode Then
    
        Set force = MakeVectorBetween(this.EndNode.displacedPosition, _
            this.StartNode.displacedPosition).WithLength(InternalForceValue)
    
    ElseIf Node Is this.EndNode Then
    
        Set force = MakeVectorBetween(this.StartNode.displacedPosition, _
            this.EndNode.displacedPosition).WithLength(InternalForceValue)
    
    Else ' unkown node
    
        Err.Raise Number:=SolutionBarElement2DErrors.UnkownNode, _
                  source:="SolutionBarElement2D.ForceInNode", _
                  Description:="Bar " & Me.ID & " does not know about node " & Node.ID
    
    End If
    
    Set ForceInNode = force
    
End Function

Public Function HasNode(ByRef Node As SolutionNode2D) As Boolean
    HasNode = (Node Is this.StartNode) Or (Node Is this.EndNode)
End Function

Public Function FinalGeometryScalingDisplacement(ByRef factor As Double) As Segment2D
    Dim segment As Segment2D
    Set segment = New Segment2D
    With segment
        Set .StartPoint = this.StartNode.DisplacedPositionScaled(factor)
        Set .EndPoint = this.EndNode.DisplacedPositionScaled(factor)
    End With
    
    Set FinalGeometryScalingDisplacement = segment
End Function

