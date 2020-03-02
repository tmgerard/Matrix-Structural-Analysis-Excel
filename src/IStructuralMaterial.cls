VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStructuralMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.Materials")
'@Interface IStructuralMaterial
Option Explicit

'@Description("Ratio of axial stress and strain during elastic deformation.")
Public Property Get ModulusOfElasticity() As Double
End Property

'@Description("The ratio of the proportional decrease in a lateral measurement to the proportional increase in length during elastic deformation.")
Public Property Get PoissonsRatio() As Double
End Property

'@Description("Ratio of shear stress and strain during elastic deformation.")
Public Property Get ShearModulus() As Double
End Property
