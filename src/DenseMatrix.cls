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

'@Folder("Model.Matrix")

Private Type TDenseMatrix
    MatrixData As IMatrixStorage
End Type

Private this As TDenseMatrix
Private Const CLASS_NAME As String = "DenseMatrix"

Private Sub Class_Initialize()
    With this
        Set .MatrixData = New IMatrixStorage
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .MatrixData = Nothing
    End With
End Sub

' IMatrix Implementation

'@Description("Returns the sum of two matrices.")
Private Function IMatrix_Add(ByVal otherMatrix As IMatrix) As IMatrix
    
    If Not this.MatrixData.Rows = otherMatrix.Rows And Not this.MatrixData.Columns = otherMatrix.Columns Then
        MatrixErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.Rows, this.MatrixData.Columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.Rows - 1
        For column = 0 To this.MatrixData.Columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) _
                + otherMatrix.Element(row, column)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.Create NewData
    
    Set IMatrix_Add = Result
    
End Function

'@Description("Returns the number of columns in matrix.")
Private Function IMatrix_Columns() As Long
    IMatrix_Columns = this.MatrixData.Columns
End Function

'@Description("Inject matrix data storage object.")
Private Sub IMatrix_Create(ByVal MatrixData As IMatrixStorage)
    Set this.MatrixData = MatrixData
End Sub

'@Description("Sets element value at specified location.")
Private Property Let IMatrix_Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    this.MatrixData.Element(rowIndex, columnIndex) = Value
End Property

'@Description("Returns element value at specified location.")
Private Property Get IMatrix_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrix_Element = this.MatrixData.Element(rowIndex, columnIndex)
End Property

'@Description("Returns true if the matrix is square.")
Private Function IMatrix_IsSquare() As Boolean
    IMatrix_IsSquare = (this.MatrixData.Rows = this.MatrixData.Columns)
End Function

'@Description("Returns the product of two matrices.")
Private Function IMatrix_Multiply(ByVal otherMatrix As IMatrix) As IMatrix

    If Not this.MatrixData.Columns = otherMatrix.Rows Then MatrixErrors.ThrowAdditionError CLASS_NAME
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.Rows, this.MatrixData.Columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    Dim InnerProduct As Long
    
    For row = 0 To this.MatrixData.Rows - 1
        For column = 0 To otherMatrix.Columns - 1
            For InnerProduct = 0 To this.MatrixData.Columns - 1
                NewData.Element(row, column) = this.MatrixData.Element(row, InnerProduct) _
                    * otherMatrix.Element(InnerProduct, column)
            Next InnerProduct
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.Create NewData
    
    Set IMatrix_Multiply = Result
    
End Function

'@Description("Returns the number of rows in matrix.")
Private Function IMatrix_Rows() As Long
    IMatrix_Rows = this.MatrixData.Rows
End Function

'@Description("Returns the product of scalar and matrix.")
Private Function IMatrix_ScalarMultiply(ByVal ScalarValue As Double) As IMatrix

    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.Rows, this.MatrixData.Columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.Rows - 1
        For column = 0 To this.MatrixData.Columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) * ScalarValue
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.Create NewData
    
    Set IMatrix_ScalarMultiply = Result
    
End Function

'@Description("Returns the difference of two matrices.")
Private Function IMatrix_Subtract(ByVal otherMatrix As IMatrix) As IMatrix

    If Not this.MatrixData.Rows = otherMatrix.Rows And Not this.MatrixData.Columns = otherMatrix.Columns Then
        MatrixErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.Rows, this.MatrixData.Columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.Rows - 1
        For column = 0 To this.MatrixData.Columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) _
                - otherMatrix.Element(row, column)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.Create NewData
    
    Set IMatrix_Subtract = Result
    
End Function

'@Description("Exchange the position of two rows within a matrix")
Private Sub IMatrix_SwapRows(ByVal FirstRowIndex As Long, ByVal SecondRowIndex As Long)

    Dim Temp As Double
    Dim column As Long
    
    For column = 0 To this.MatrixData.Columns - 1
        With this.MatrixData
            Temp = .Element(SecondRowIndex, column)
            .Element(SecondRowIndex, column) = .Element(FirstRowIndex, column)
            .Element(FirstRowIndex, column) = Temp
        End With
    Next column

End Sub

'@Description("Return the transposes of a matrix")
Private Function IMatrix_Transpose() As IMatrix

    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.Rows, this.MatrixData.Columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To this.MatrixData.Columns - 1
        For column = 0 To this.MatrixData.Rows - 1
            NewData.Element(row, column) = this.MatrixData(column, row)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.Create NewData
    
    Set IMatrix_Transpose = Result
    
End Function

