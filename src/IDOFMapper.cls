VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDOFMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Property Get GetDofEq(ByRef elementIndex As Long, ByRef nodeIndex As Long) As Long
End Property

Public Property Get DofMap() As Matrix
End Property
