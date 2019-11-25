VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LupSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Solver.Direct Methods")
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
    
    Dim y As IVector
    Set y = ForwardSubstitution(coefficientMatrix, solutionVector)
    
    Dim result As IVector
    Set result = BackSubstitution(coefficientMatrix, y)
    
    Set Solve = result

End Function

Private Function BackSubstitution(ByVal factorizedMatrix As IMatrix, ByVal forwardSolvedVector As IVector) As IVector
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactory(DenseColumnVector)
    
    Dim result As New IVector
    Set result = New DenseVector
    result.Storage = factory.Create(forwardSolvedVector.length)
    
    Set factory = Nothing
    
    Dim i As Long
    Dim j As Long
    Dim sum As Double
    For i = forwardSolvedVector.length - 1 To 0 Step -1
        sum = 0
        For j = 0 To forwardSolvedVector.length - 1
            sum = sum + factorizedMatrix.Element(i, j) * result.Element(j)
        Next j
        
        result.Element(i) = (forwardSolvedVector.Element(i) - sum) _
            / factorizedMatrix.Element(i, i)
    Next i
    
    Set BackSubstitution = result
    
End Function

Private Function ForwardSubstitution(ByVal factorizedMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactory(DenseColumnVector)
    
    Dim y As New IVector
    Set y = New DenseVector
    y.Storage = factory.Create(solutionVector.length)
    
    Set factory = Nothing
    
    Dim i As Long
    Dim j As Long
    Dim sum As Double
    For i = 0 To solutionVector.length - 1
        sum = 0
        For j = 0 To i
            sum = sum + factorizedMatrix.Element(i, j) * y.Element(j)
        Next j
        
        y.Element(i) = solutionVector.Element(this.permutation.Element(i)) - sum
    Next i
    
    Set ForwardSubstitution = y
    
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
            
            If Math.Abs(coefficientMatrix.Element(rowIndex, pivotRowIndex)) > pivotValue Then
                pivotValue = Math.Abs(coefficientMatrix.Element(rowIndex, pivotRowIndex))
                swapRowIndex = rowIndex
            End If
        
        Next rowIndex
        
        ' if pivot value is zero then the matrix is singular and no solution exists
        If pivotValue = 0 Then SolverErrors.ThrowSingularMatrixError CLASS_NAME
        
        this.permutation.SwapValues pivotRowIndex, swapRowIndex
        coefficientMatrix.SwapRows pivotRowIndex, swapRowIndex
        
        ' reduce column below pivot
        For rowIndex = pivotRowIndex + 1 To numRows - 1
        
            With coefficientMatrix
                .Element(rowIndex, pivotRowIndex) = .Element(rowIndex, pivotRowIndex) _
                    / .Element(pivotRowIndex, pivotRowIndex)
            End With
            
            ' update lower submatrix with Schur complement
            For columnIndex = pivotRowIndex + 1 To numRows - 1
            
                With coefficientMatrix
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
    
    this.permutation.Storage = factory.Create(rowCount)
    
    Set factory = Nothing
    
    ' store row indices in vector
    Dim i As Long
    For i = 0 To rowCount - 1
        this.permutation.Element(i) = i
    Next i

End Sub

' ILinearSolver Implementation
Private Function ILinearSolver_Solve(ByVal coefficientMatrix As IMatrix, ByVal solutionVector As IVector) As IVector
    Set ILinearSolver_Solve = Me.Solve(coefficientMatrix, solutionVector)
End Function
