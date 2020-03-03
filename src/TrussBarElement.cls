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

    

End Function
