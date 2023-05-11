VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BarElement2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit
Implements ILinearElement

Private Type TBarElement
    BarID As Long
    StartNode As Node2D
    EndNode As Node2D
    barArea As Double                            ' Note: A materials class (interface?) might be better for implementing beam elements
    ElasticModulus As Double
End Type

Private this As TBarElement

Private Sub Class_Initialize()
    With this
        Set .StartNode = New Node2D
        Set .EndNode = New Node2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .StartNode = Nothing
        Set .EndNode = Nothing
    End With
End Sub

Public Property Get ID() As Long
    ID = this.BarID
End Property

Public Property Let ID(ByRef BarID As Long)
    this.BarID = BarID
End Property

Public Property Get StartNode() As Node2D
    Set StartNode = this.StartNode
End Property

Public Property Set StartNode(ByRef Node As Node2D)
    Set this.StartNode = Node
End Property

Public Property Get EndNode() As Node2D
    Set EndNode = this.EndNode
End Property

Public Property Set EndNode(ByRef Node As Node2D)
    Set this.EndNode = Node
End Property

Public Property Get Area() As Double
    Area = this.barArea
End Property

Public Property Let Area(ByRef barArea As Double)
    this.barArea = barArea
End Property

Public Property Get ElasticModulus() As Double
    ElasticModulus = this.ElasticModulus
End Property

Public Property Let ElasticModulus(ByRef Modulus_E As Double)
    this.ElasticModulus = Modulus_E
End Property

Public Function ElementNodes() As Collection
    Dim result As Collection
    Set result = New Collection
    
    result.Add this.StartNode
    result.Add this.EndNode
    
    Set ElementNodes = result
End Function

Public Function Geometry() As Segment2D
    Dim barSegment As Segment2D
    Set barSegment = New Segment2D
    With barSegment
        Set .StartPoint = this.StartNode.Position
        Set .EndPoint = this.EndNode.Position
    End With
    
    Set Geometry = barSegment
End Function

Public Function Length() As Double
    Length = Me.Geometry.Length
End Function

Public Function K_Global() As Matrix

    Dim direction As Vector2D
    Set direction = Me.Geometry.DirectionVector
    
    Dim axialStiffness  As Double
    axialStiffness = Me.ElasticModulus * Me.Area / Me.Length
    
    Dim c As Double
    c = direction.Cosine
    
    Dim s As Double
    s = direction.Sine
    
    Dim c2_axialStiffness As Double
    c2_axialStiffness = (c ^ 2) * axialStiffness
    
    Dim s2_axialStiffness As Double
    s2_axialStiffness = (s ^ 2) * axialStiffness
    
    Dim sc_axialStiffness As Double
    sc_axialStiffness = (s * c) * axialStiffness
    
    Dim stiffMatrixData(15) As Double
    stiffMatrixData(0) = c2_axialStiffness
    stiffMatrixData(1) = sc_axialStiffness
    stiffMatrixData(2) = -c2_axialStiffness
    stiffMatrixData(3) = -sc_axialStiffness
    stiffMatrixData(4) = sc_axialStiffness
    stiffMatrixData(5) = s2_axialStiffness
    stiffMatrixData(6) = -sc_axialStiffness
    stiffMatrixData(7) = -s2_axialStiffness
    stiffMatrixData(8) = -c2_axialStiffness
    stiffMatrixData(9) = -sc_axialStiffness
    stiffMatrixData(10) = c2_axialStiffness
    stiffMatrixData(11) = sc_axialStiffness
    stiffMatrixData(12) = -sc_axialStiffness
    stiffMatrixData(13) = -s2_axialStiffness
    stiffMatrixData(14) = sc_axialStiffness
    stiffMatrixData(15) = s2_axialStiffness
    
    Dim stiffnessMatrix As Matrix
    Set stiffnessMatrix = New Matrix
    Set stiffnessMatrix = stiffnessMatrix.SetSize(4, 4).SetData(stiffMatrixData)
    
    Set GlobalStiffnessMatrix = stiffnessMatrix

End Function

' ILinearElement Implementation

Private Property Get ILinearElement_EndNode() As Node2D
    Set ILinearElement_EndNode = Me.EndNode
End Property

Private Function ILinearElement_GlobalKMatrix() As Matrix
    Set ILinearElement_GlobalKMatrix = Me.K_Global
End Function

Private Property Let ILinearElement_ID(ByRef value As Long)
    Me.ID = value
End Property

Private Property Get ILinearElement_ID() As Long
    ILinearElement_ID = Me.ID
End Property

Private Function ILinearElement_LocalKMatrix() As Matrix
    Set ILinearElement_LocalKMatrix = Nothing
End Function

Private Property Get ILinearElement_StartNode() As Node2D
    Set ILinearElement_StartNode = Me.StartNode
End Property
