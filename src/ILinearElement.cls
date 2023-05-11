VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinearElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit
' Interface for linear elements (bar and beam) ensuring
' access to fields required for basic operations and
' stiffness matrix assembly

Public Property Get ID() As Long
End Property

Public Property Get StartNode() As Node2D
End Property

Public Property Get EndNode() As Node2D
End Property

Public Function K_Local() As Matrix
End Function

Public Function K_Global() As Matrix
End Function
