VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SingluarCeofficientMatrixAStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.LinearAlgebra.SolverStubs")
'@IgnoreModule
Option Explicit
Implements IMatrix
Implements ICloneable

Private matrixStub As IMatrix
Const ROW_COUNT As Long = 3
Const COLUMN_COUNT As Long = 3

Private Sub Class_Initialize()
    Set matrixStub = New DenseMatrix
    With matrixStub
        Set .Storage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(ROW_COUNT, COLUMN_COUNT)
        .Storage.Element(0, 0) = 0
        .Storage.Element(0, 1) = 2
        .Storage.Element(0, 2) = 1
        .Storage.Element(1, 0) = 0
        .Storage.Element(1, 1) = 1
        .Storage.Element(1, 2) = -1
        .Storage.Element(2, 0) = 0
        .Storage.Element(2, 1) = 0
        .Storage.Element(2, 2) = 2
    End With
End Sub

Private Sub Class_Terminate()
    Set matrixStub = Nothing
End Sub

Private Function ICloneable_Clone() As Object
    Dim cloned As ICloneable
    Set cloned = matrixStub
    Set ICloneable_Clone = cloned.Clone
End Function

Private Function IMatrix_Add(ByVal otherMatrix As IMatrix) As IMatrix

End Function

Private Function IMatrix_Columns() As Long
    IMatrix_Columns = matrixStub.Columns
End Function

Private Function IMatrix_Equals(ByVal otherMatrix As IMatrix) As Boolean

End Function

Private Function IMatrix_IsSquare() As Boolean
    IMatrix_IsSquare = matrixStub.IsSquare
End Function

Private Function IMatrix_Multiply(ByVal otherMatrix As IMatrix) As IMatrix

End Function

Private Function IMatrix_Rows() As Long
    IMatrix_Rows = matrixStub.Rows
End Function

Private Function IMatrix_ScalarMultiply(ByVal scalarValue As Double) As IMatrix

End Function

Private Sub IMatrix_SetSize(ByVal Rows As Long, ByVal Columns As Long)

End Sub

Private Property Set IMatrix_storage(ByVal RHS As IMatrixStorage)

End Property

Private Property Get IMatrix_storage() As IMatrixStorage
    Set IMatrix_storage = matrixStub.Storage
End Property

Private Function IMatrix_Subtract(ByVal otherMatrix As IMatrix) As IMatrix

End Function

Private Sub IMatrix_SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)
    matrixStub.SwapRows firstRowIndex, secondRowIndex
End Sub

Private Function IMatrix_Transpose() As IMatrix

End Function
