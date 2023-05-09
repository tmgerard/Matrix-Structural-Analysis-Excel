VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BeamElement2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TBeamElement2D
    BeamID As Long
    Material As IPlaneBeamMaterial
    Section As IPlaneBeamCrossSection
    StartNode As Node2D
    EndNode As Node2D
End Type

Private this As TBeamElement2D

Private Sub Class_Initialize()
    With this
        .StartNode = New Node2D
        .EndNode = New Node2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        .StartNode = Nothing
        .EndNode = Nothing
    End With
End Sub

Public Property Get ID() As Long
    ID = this.BeamID
End Property

Public Property Let ID(ByRef BeamID As Long)
    this.BeamID = BeamID
End Property

Public Property Get StartNode() As Node2D
    Set StartNode = this.StartNode
End Property

Public Property Set StartNode(ByRef node As Node2D)
    Set this.StartNode = node
End Property

Public Property Get EndNode() As Node2D
    Set EndNode = this.EndNode
End Property

Public Property Set EndNode(ByRef node As Node2D)
    Set this.EndNode = node
End Property

Public Property Get Material() As IPlaneBeamMaterial
    Set Material = this.Material
End Property

Public Property Set Material(ByRef BeamMaterial As IPlaneBeamMaterial)
    Set this.Material = BeamMaterial
End Property

Public Property Get Section() As IPlaneBeamCrossSection
    Set Section = this.Section
End Property

Public Property Set Section(ByRef BeamSection As IPlaneBeamCrossSection)
    Set this.Section = BeamSection
End Property

Public Function GetSegment2D() As Segment2D
    Dim beamSegment As Segment2D
    Set beamSegment = New Segment2D
    With beamSegment
        Set .StartPoint = this.StartNode.Position
        Set .EndPoint = this.EndNode.Position
    End With
    
    Set GetSegment2D = beamSegment
End Function

Public Function Length() As Double
    Length = Me.GetSegment2D.Length
End Function

Public Function K_Local() As Matrix
    ' Define the unique stiffness terms
    Dim A As Double, E As Double, I As Double, L As Double
    A = this.Section.Area
    E = this.Material.E
    I = this.Section.I
    L = Me.Length

    Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double
    k1 = 4 * E * I / L
    k2 = 2 * E * I / L ^ 2
    k3 = 6 * E * I / L ^ 2
    k4 = 12 * E * I / L ^ 3
    k5 = A * E / L
    
    Dim matrixData(35) As Double
    ' row 1
    matrixData(0) = k1
    matrixData(1) = k2
    matrixData(2) = k3
    matrixData(3) = -k3
    matrixData(4) = 0
    matrixData(5) = 0
    ' row 2
    matrixData(6) = k2
    matrixData(7) = k1
    matrixData(8) = k3
    matrixData(9) = -k3
    matrixData(10) = 0
    matrixData(11) = 0
    ' row 3
    matrixData(12) = k3
    matrixData(13) = k3
    matrixData(14) = k4
    matrixData(15) = -k4
    matrixData(16) = 0
    matrixData(17) = 0
    ' row 4
    matrixData(18) = -k3
    matrixData(19) = -k3
    matrixData(20) = -k4
    matrixData(21) = k4
    matrixData(22) = 0
    matrixData(23) = 0
    ' row 5
    matrixData(24) = 0
    matrixData(25) = 0
    matrixData(26) = 0
    matrixData(27) = 0
    matrixData(28) = k5
    matrixData(29) = -k5
    ' row 6
    matrixData(30) = 0
    matrixData(31) = 0
    matrixData(32) = 0
    matrixData(33) = 0
    matrixData(34) = -k5
    matrixData(35) = k5
    
    Dim kMatrix As Matrix
    Set kMatrix = New Matrix
    Set kMatrix = kMatrix.SetSize(6, 6).SetData(matrixData)
    
    Set K_Local = kMatrix
End Function





