VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LupSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Solver.DirectMethods")
'@IgnoreModule UseMeaningfulName, ObjectVariableNotSet
Option Explicit
Implements ILinearSolver

Private Type TLupSolver
    permutation As IVector
End Type

Private this As TLupSolver
Private Const CLASS_NAME As String = "DenseMatrix"

Private Sub Class_Initialize()
    With this
        Set .permutation = New DenseVector
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .permutation = Nothing
    End With
End Sub

Public Function Solve(ByVal coefficientMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
    
    If Not coefficientMatrix.IsSquare Then SolverErrors.ThrowDecompositionError CLASS_NAME
    
    LupDecomposition coefficientMatrix
    
    '@Ignore UseMeaningfulName
    Dim Y As IVector
    Set Y = ForwardSubstitution(coefficientMatrix, solutionVector)
    
    Dim result As IVector
    Set result = BackSubstitution(coefficientMatrix, Y)
    
    Set Solve = result

End Function

Private Function BackSubstitution(ByVal factorizedMatrix As IMatrix, ByVal forwardSolvedVector As IVector) As IVector
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactory(DenseColumnVector)
    
    Dim result As IVector
    Set result = New DenseVector
    Set result.Storage = factory.Create(forwardSolvedVector.length)
    
    Set factory = Nothing
    
    Dim I As Long
    Dim J As Long
    Dim sum As Double
    For I = forwardSolvedVector.length - 1 To 0 Step -1
        sum = 0
        For J = 0 To forwardSolvedVector.length - 1
            sum = sum + factorizedMatrix.Storage.Element(I, J) * result.Storage.Element(J)
        Next J
        
        result.Storage.Element(I) = (forwardSolvedVector.Storage.Element(I) - sum) _
            / factorizedMatrix.Storage.Element(I, I)
    Next I
    
    Set BackSubstitution = result
    
End Function

Private Function ForwardSubstitution(ByVal factorizedMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactory(DenseColumnVector)
    
    Dim Y As IVector
    Set Y = New DenseVector
    Set Y.Storage = factory.Create(solutionVector.length)
    
    Set factory = Nothing
    
    Dim I As Long
    Dim J As Long
    Dim sum As Double
    For I = 0 To solutionVector.length - 1
        sum = 0
        For J = 0 To I
            sum = sum + factorizedMatrix.Storage.Element(I, J) * Y.Storage.Element(J)
        Next J
        
        Y.Storage.Element(I) = solutionVector.Storage.Element(this.permutation.Storage.Element(I)) - sum
    Next I
    
    Set ForwardSubstitution = Y
    
End Function

Private Sub LupDecomposition(ByRef coefficientMatrix As IMatrix)

    Dim numRows As Long
    numRows = coefficientMatrix.Rows

    InitializePermutationVector rowCount:=numRows
    
    Dim rowIndex As Long
    Dim pivotRowIndex As Long
    Dim swapRowIndex As Long
    Dim columnIndex As Long
    Dim pivotValue As Double
    
    For pivotRowIndex = 0 To numRows - 1
    
        pivotValue = 0
        
        ' find largest value in column to pivot on
        For rowIndex = pivotRowIndex To numRows - 1
            
            If Math.Abs(coefficientMatrix.Storage.Element(rowIndex, pivotRowIndex)) > pivotValue Then
                pivotValue = Math.Abs(coefficientMatrix.Storage.Element(rowIndex, pivotRowIndex))
                swapRowIndex = rowIndex
            End If
        
        Next rowIndex
        
        ' if pivot value is zero then the matrix is singular and no solution exists
        If pivotValue = 0 Then SolverErrors.ThrowSingularMatrixError CLASS_NAME
        
        this.permutation.SwapValues pivotRowIndex, swapRowIndex
        coefficientMatrix.SwapRows pivotRowIndex, swapRowIndex
        
        ' reduce column below pivot
        For rowIndex = pivotRowIndex + 1 To numRows - 1
        
            With coefficientMatrix.Storage
                .Element(rowIndex, pivotRowIndex) = .Element(rowIndex, pivotRowIndex) _
                    / .Element(pivotRowIndex, pivotRowIndex)
            End With
            
            ' update lower submatrix with Schur complement
            For columnIndex = pivotRowIndex + 1 To numRows - 1
            
                With coefficientMatrix.Storage
                    .Element(rowIndex, columnIndex) = .Element(rowIndex, columnIndex) _
                        - .Element(rowIndex, pivotRowIndex) * .Element(pivotRowIndex, columnIndex)
                End With
                
            Next columnIndex
        
        Next rowIndex
    
    Next pivotRowIndex

End Sub

Private Sub InitializePermutationVector(ByVal rowCount As Long)

    ' set up the permutation vector storage size
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactory(DenseColumnVector)
    
    Set this.permutation.Storage = factory.Create(rowCount)
    
    Set factory = Nothing
    
    ' store row indices in vector
    Dim I As Long
    For I = 0 To rowCount - 1
        this.permutation.Storage.Element(I) = I
    Next I

End Sub

' ILinearSolver Implementation
Private Function ILinearSolver_Solve(ByVal coefficientMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
    Set ILinearSolver_Solve = Me.Solve(coefficientMatrix, solutionVector)
End Function




