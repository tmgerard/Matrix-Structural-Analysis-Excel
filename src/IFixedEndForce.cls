VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFixedEndForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.FixedEndForces")
Option Explicit

Public Enum FixedEndForceLocation
    Left
    Right
End Enum

Public Function FixedEndMoment(ByVal location As FixedEndForceLocation) As Double
End Function

Public Function FixedEndReaction(ByVal location As FixedEndForceLocation) As Double
End Function
