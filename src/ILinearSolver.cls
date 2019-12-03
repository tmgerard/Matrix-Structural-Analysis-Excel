VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinearSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.LinearAlgebra.Solver.Interface")
'@Interface ILinearSolver
Option Explicit

Public Function Solve(ByVal coefficientMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
End Function
