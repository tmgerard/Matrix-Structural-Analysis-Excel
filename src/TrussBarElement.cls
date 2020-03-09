VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrussBarElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.Elements")
Option Explicit
Implements IElement

Private Type tTrussBarElement
    StartNode As INode
    EndNode As INode
    CrossSection As IStructuralCrossSection
    Material As IStructuralMaterial
End Type

Private this As tTrussBarElement
Private Const LocalDOFs As Integer = 2
Private Const GlobalDOFs As Integer = 4

Public Property Get StartNode() As INode
    StartNode = this.StartNode
End Property

Public Property Set StartNode(ByRef value As INode)
    Set this.StartNode = value
    ' TODO: release rotational degrees of freedom
End Property

Public Property Get EndNode() As INode
    EndNode = this.EndNode
End Property

Public Property Set EndNode(ByRef value As INode)
    Set this.EndNode = value
    ' TODO: release rotational degrees of freedom
End Property

Public Property Get CrossSection() As INode
    CrossSection = this.CrossSection
End Property

Public Property Set CrossSection(ByRef value As INode)
    Set this.CrossSection = value
End Property

Public Property Get Material() As INode
    Material = this.Material
End Property

Public Property Set Material(ByRef value As INode)
    Set this.Material = Material
End Property

Public Function ElementStiffnessMatrix() As DenseMatrix

    Dim stiffnessMatrix As DenseMatrix
    Set stiffnessMatrix = New DenseMatrix
    
    With stiffnessMatrix
        Set .Storage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(LocalDOFs, LocalDOFs)
        
        Dim axialStiffness As Double
        axialStiffness = this.CrossSection.Area * this.Material.ModulusOfElasticity / _
            this.StartNode.Coordinate.Distance(this.EndNode.Coordinate)
        
        .Storage.Element(0, 0) = axialStiffness
        .Storage.Element(0, 1) = -axialStiffness
        .Storage.Element(1, 0) = axialStiffness
        .Storage.Element(1, 1) = -axialStiffness
        
    End With
    
    Set ElementStiffnessMatrix = stiffnessMatrix

End Function
