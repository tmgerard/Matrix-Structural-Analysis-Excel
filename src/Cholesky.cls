VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cholesky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Equations")
Option Explicit
Implements ILinearSolver

Public Enum CholeskyErrors
    BadSize = vbObjectError + 600
    NotSquare
    Unsolvable
End Enum

'@Description "Solves symmetric, positive definite system of linear equations."
Public Function Solve(ByRef mat As Matrix, ByRef vec As Vector) As Vector
Attribute Solve.VB_Description = "Solves symmetric, positive definite system of linear equations."
    ' check that matrix and vector work for Cholesky Decomposition
    ValidateSystem mat, vec
    
    Dim lowMat As Matrix
    Set lowMat = LowDecomposition(mat)
    
    Dim lowSolution As Vector
    Set lowSolution = ForwardSubstitution(lowMat, vec)
    
    Set Solve = BackSubstutution(lowMat, lowSolution)
    
End Function

'@Description "Checks that the coefficient matrix and free vector are compatible with Cholesky Decomposition."
Private Sub ValidateSystem(ByRef mat As Matrix, ByRef vec As Vector)
Attribute ValidateSystem.VB_Description = "Checks that the coefficient matrix and free vector are compatible with Cholesky Decomposition."
    If Not mat.Columns = vec.Length Then
        Err.Raise Number:=CholeskyErrors.BadSize, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="Size mismatch between matrix and vector."
    End If
    
    If Not mat.IsSquare Then
        Err.Raise Number:=CholeskyErrors.NotSquare, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="System matrix must be square"
    End If
End Sub

'@Description "Cholesky Decomposition of the lower triangular matrix."
Public Function LowDecomposition(ByRef mat As Matrix) As Matrix
Attribute LowDecomposition.VB_Description = "Cholesky Decomposition of the lower triangular matrix."

    On Error GoTo ErrorHandler
    
    Dim size As Long
    size = mat.Rows
    
    Dim decomp As Matrix
    Set decomp = New Matrix
    Set decomp = decomp.SetSize(size, size)
    
    Dim I As Long
    Dim j As Long
    Dim k As Long
    Dim sqSum As Double
    Dim m_ij As Double
    Dim diagValue As Double
    Dim nonDiagValue
    Dim nonDiagSum As Double
    Dim l_ik As Double
    Dim l_jk As Double
    Dim l_jj As Double
    
    For I = 0 To size - 1
    
        sqSum = 0
        For j = 0 To I
            
            m_ij = mat.ValueAt(I, j)
            If I = j Then ' value on main diagonal
                diagValue = Math.Sqr(m_ij - sqSum)
                decomp.ValueAt(I, j) = diagValue
            Else ' value under main diagonal
                nonDiagSum = 0
                
                For k = 0 To j - 1
                    
                    l_ik = decomp.ValueAt(I, k)
                    l_jk = decomp.ValueAt(j, k)
                    nonDiagSum = nonDiagSum + l_ik * l_jk
                    
                Next k
                
                l_jj = decomp.ValueAt(j, j)
                nonDiagValue = (m_ij - nonDiagSum) / l_jj
                sqSum = sqSum + nonDiagValue * nonDiagValue
                
                decomp.ValueAt(I, j) = nonDiagValue
   
            End If
            
        Next j
        
    Next I
    
    Set LowDecomposition = decomp
    
    Exit Function
    
ErrorHandler:
    
    Err.Raise CholeskyErrors.Unsolvable, "Cholesky.LowDecomposition", _
        "Global stiffness matrix not solvable by Choleski Decomposition. Please verify the structure inputs."
    
End Function

'@Description "Obtains solution vector for lower triangular matrix."
Public Function ForwardSubstitution(ByRef mat As Matrix, ByRef vec As Vector) As Vector
Attribute ForwardSubstitution.VB_Description = "Obtains solution vector for lower triangular matrix."

    Dim size As Long
    size = vec.Length
    
    Dim solution As Vector
    Set solution = New Vector
    Set solution = solution.SetLength(size)
    
    Dim I As Long
    Dim j As Long
    Dim l_ij As Double
    Dim y_j As Double
    Dim b_i As Double
    Dim l_ii As Double
    Dim sum As Double
    Dim solutionValue As Double
    
    For I = 0 To size - 1
    
        sum = 0
        
        For j = 0 To I - 1
        
            l_ij = mat.ValueAt(I, j)
            y_j = solution.ValueAt(j)
            sum = sum + l_ij * y_j
        
        Next j
        
        b_i = vec.ValueAt(I)
        l_ii = mat.ValueAt(I, I)
        solutionValue = (b_i - sum) / l_ii
        solution.ValueAt(I) = solutionValue
        
    Next I
    
    Set ForwardSubstitution = solution

End Function

'@Description "Obtains solution vector of the upper triangular matrix. The upeer triangluar matrix is the transpose of the lower triangluar matrix."
Public Function BackSubstutution(ByRef mat As Matrix, ByRef vec As Vector) As Vector
Attribute BackSubstutution.VB_Description = "Obtains solution vector of the upper triangular matrix. The upeer triangluar matrix is the transpose of the lower triangluar matrix."
    
    Dim size As Long
    size = vec.Length
    
    Dim lastIndex As Long
    lastIndex = size - 1
    
    Dim solution As Vector
    Set solution = New Vector
    Set solution = solution.SetLength(size)
    
    Dim I As Long
    Dim j As Long
    Dim sum As Double
    Dim u_ij As Double
    Dim x_j As Double
    Dim y_i As Double
    Dim u_ii As Double
    Dim solutionValue As Double
    
    For I = lastIndex To 0 Step -1
        
        sum = 0
        
        For j = I + 1 To size - 1
        
            u_ij = mat.ValueAtTransposed(I, j)
            x_j = solution.ValueAt(j)
            sum = sum + u_ij * x_j
        
        Next j
        
        y_i = vec.ValueAt(I)
        u_ii = mat.ValueAtTransposed(I, I)
        solutionValue = (y_i - sum) / u_ii
        solution.ValueAt(I) = solutionValue
        
    Next I
    
    Set BackSubstutution = solution
    
End Function

' ILinearSolver Implementation
Private Function ILinearSolver_Solve(mat As Matrix, vec As Vector) As Vector
    Set ILinearSolver_Solve = Me.Solve(mat, vec)
End Function
