VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinearSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("StructuralAnalysis.Equations")
Option Explicit

Public Function Solve(ByRef mat As Matrix, ByRef vec As Vector) As Vector
End Function
