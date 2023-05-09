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
Private Const MaxDofs As Long = 6

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
    Dim A As Double, E As Double, i As Double, L As Double
    A = this.Section.Area
    E = this.Material.E
    i = this.Section.i
    L = Me.Length

    Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double
    k1 = A * E / L
    k2 = 12 * E * i / L ^ 3
    k3 = 6 * E * i / L ^ 2
    k4 = 4 * E * i / L
    k5 = 2 * E * i / L ^ 2
    
    Dim kMatrix As Matrix
    Set kMatrix = New Matrix
    Set kMatrix = kMatrix.SetSize(6, 6)
    
    With kMatrix
        .ValueAt(0, 0) = k1
        .ValueAt(0, 3) = -k1
        .ValueAt(1, 1) = k2
        .ValueAt(1, 2) = k3
        .ValueAt(1, 4) = -k2
        .ValueAt(1, 5) = -k3
        .ValueAt(2, 1) = k3
        .ValueAt(2, 2) = k4
        .ValueAt(2, 4) = -k3
        .ValueAt(2, 4) = k5
        .ValueAt(3, 0) = -k1
        .ValueAt(3, 3) = k1
        .ValueAt(4, 1) = -k2
        .ValueAt(4, 2) = -k3
        .ValueAt(4, 4) = k2
        .ValueAt(4, 5) = -k3
        .ValueAt(5, 1) = k3
        .ValueAt(5, 2) = k5
        .ValueAt(5, 4) = -k3
        .ValueAt(5, 5) = k4
    End With
    
    Set K_Local = kMatrix
End Function

Public Function K_Global() As Matrix
    Dim T As Matrix
    Set T = GetTransform
    
    Dim temp As Matrix
    Set temp = T.Transpose  ' TODO: implement matrix multiplication
End Function

Private Function GetTransform() As Matrix
    
    Dim vec As Vector2D
    Set vec = MakeVector(1, 0)  ' unit vector along x-axis
    
    Dim angle As Double
    angle = vec.AngleTo(this.EndNode.Position - this.StartNode.Position)
    
    Dim transform As Matrix
    Set transform = New Matrix
    Set transform = transform.SetSize(MaxDofs, MaxDofs)
    
    Dim c As Double
    c = Math.cos(angle)
    
    ' for rounding errors
    If Doubles.IsEffectivelyZero(c) Then c = 0#
    If Doubles.IsEffectivelyOne(c) Then c = 1#
    
    Dim s As Double
    s = Math.sin(angle)
    
    ' for rounding errors
    If Doubles.IsEffectivelyZero(s) Then s = 0#
    If Doubles.IsEffectivelyOne(s) Then s = 1#
    
    With transform
        .ValueAt(0, 0) = c
        .ValueAt(0, 1) = s
        .ValueAt(1, 0) = -s
        .ValueAt(1, 1) = c
        .ValueAt(2, 2) = 1#
        .ValueAt(3, 3) = c
        .ValueAt(3, 4) = s
        .ValueAt(4, 3) = -s
        .ValueAt(4, 4) = c
        .ValueAt(5, 5) = 1#
    End With
    
    Set GetTransform = transform
    
End Function




