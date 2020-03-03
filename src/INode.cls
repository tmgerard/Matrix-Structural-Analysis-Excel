VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.Node")
'@Interface INode
Option Explicit

Public Enum SupportCondition
    Roller = 0
    Pinned
    Slider
    Fixed
End Enum

Public Enum DegreeOfFreedom
    Fx = 0
    Mx
    Fy
    My
    Fz
    Mz
End Enum

