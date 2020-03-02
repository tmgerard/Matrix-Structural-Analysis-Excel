VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolverErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Solver.Errors")

Option Explicit

Public Enum SolverError
    SingularMatrix = vbObjectError + 300
    Decomposition
End Enum

Public Sub ThrowSingularMatrixError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Matrix is singular. A real solution does not exist."
    
    Err.Raise Number:=SolverError.SingularMatrix, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowDecompositionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Matrix cannot be factorized by decomposition method."
    
    Err.Raise Number:=SolverError.Decomposition, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub


