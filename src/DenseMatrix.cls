VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMatrix
Implements ICloneable

'@Folder("StructuralAnalysis.Math.LinearAlgebra.Matrix")
'@IgnoreModule ObjectVariableNotSet

Private Type TDenseMatrix
    matrixData As IMatrixStorage
    operator As MatrixOperator
End Type

Private this As TDenseMatrix

Private Sub Class_Initialize()
    With this
        Set .matrixData = New IMatrixStorage
        Set .operator = New MatrixOperator
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .matrixData = Nothing
        Set .operator = Nothing
    End With
End Sub

Public Function Add(ByVal otherMatrix As IMatrix) As DenseMatrix
    Set Add = this.operator.Add(Me, otherMatrix)
End Function

Public Function Columns() As Long
    Columns = this.matrixData.Columns
End Function

Public Function Clone() As DenseMatrix
    Dim matrixData As ICloneable
    Set matrixData = this.matrixData
    
    Dim cloneMatrix As DenseMatrix
    Set cloneMatrix = New DenseMatrix
    
    Set cloneMatrix.Storage = matrixData.Clone
    
    Set Clone = cloneMatrix
End Function

Public Property Set Storage(ByVal matrixData As IMatrixStorage)
    Set this.matrixData = matrixData
End Property

Public Property Get Storage() As IMatrixStorage
    Set Storage = this.matrixData
End Property

Public Function Equals(ByVal otherMatrix As IMatrix) As Boolean
    Equals = this.operator.Equals(Me, otherMatrix)
End Function

Public Function IsSquare() As Boolean
    IsSquare = (this.matrixData.Rows = this.matrixData.Columns)
End Function

Public Function Multiply(ByVal otherMatrix As IMatrix) As DenseMatrix
    Set Multiply = this.operator.Multiply(Me, otherMatrix)
End Function

Public Function ScalarMultiply(ByVal scalarValue As Double) As DenseMatrix
    Set ScalarMultiply = this.operator.ScalarMultiply(Me, scalarValue)
End Function

Public Sub SetSize(ByVal Rows As Long, ByVal Columns As Long)
    this.matrixData.SetSize Rows, Columns
End Sub

Public Function Subtract(ByVal otherMatrix As IMatrix) As DenseMatrix
    Set Subtract = this.operator.Subtract(Me, otherMatrix)
End Function

Public Function Rows() As Long
    Rows = this.matrixData.Rows
End Function

Public Sub SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)
    this.operator.SwapRows Me, firstRowIndex, secondRowIndex
End Sub

Public Function Transpose() As DenseMatrix
    Set Transpose = this.operator.Transpose(Me)
End Function

' ICloneable Implementation

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Me.Clone
End Function

' IMatrix Implementation

'@Description("Returns the sum of two matrices.")
Private Function IMatrix_Add(ByVal otherMatrix As IMatrix) As IMatrix
    Set IMatrix_Add = Me.Add(otherMatrix)
End Function

'@Description("Returns the number of columns in matrix.")
Private Function IMatrix_Columns() As Long
    IMatrix_Columns = Me.Columns
End Function

Private Function IMatrix_Equals(ByVal otherMatrix As IMatrix) As Boolean
    IMatrix_Equals = Me.Equals(otherMatrix)
End Function

Private Sub IMatrix_SetSize(ByVal Rows As Long, ByVal Columns As Long)
    Me.SetSize Rows, Columns
End Sub

Private Property Set IMatrix_storage(ByVal Storage As IMatrixStorage)
    Set Me.Storage = Storage
End Property

Private Property Get IMatrix_storage() As IMatrixStorage
    Set IMatrix_storage = Me.Storage
End Property

'@Description("Returns true if the matrix is square.")
Private Function IMatrix_IsSquare() As Boolean
    IMatrix_IsSquare = Me.IsSquare
End Function

'@Description("Returns the product of two matrices.")
Private Function IMatrix_Multiply(ByVal otherMatrix As IMatrix) As IMatrix
    Set IMatrix_Multiply = Me.Multiply(otherMatrix)
End Function

'@Description("Returns the number of rows in matrix.")
Private Function IMatrix_Rows() As Long
    IMatrix_Rows = Me.Rows
End Function

'@Description("Returns the product of scalar and matrix.")
Private Function IMatrix_ScalarMultiply(ByVal scalarValue As Double) As IMatrix
    Set IMatrix_ScalarMultiply = Me.ScalarMultiply(scalarValue)
End Function

'@Description("Returns the difference of two matrices.")
Private Function IMatrix_Subtract(ByVal otherMatrix As IMatrix) As IMatrix
    Set IMatrix_Subtract = Me.Subtract(otherMatrix)
End Function

'@Description("Exchange the position of two rows within a matrix")
Private Sub IMatrix_SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)
    Me.SwapRows firstRowIndex, secondRowIndex
End Sub

'@Description("Return the transposes of a matrix")
Private Function IMatrix_Transpose() As IMatrix
    Set IMatrix_Transpose = Me.Transpose
End Function

