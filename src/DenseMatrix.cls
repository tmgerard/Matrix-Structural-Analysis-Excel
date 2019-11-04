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

' TODO: Implement a vector multiply function

'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix")

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

Public Function Add(ByVal otherMatrix As IMatrix) As DenseMatrix

    If Not this.MatrixData.rows = otherMatrix.rows Or Not this.MatrixData.columns = otherMatrix.columns Then
        MatrixErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.rows, this.MatrixData.columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.rows - 1
        For column = 0 To this.MatrixData.columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) _
                + otherMatrix.Element(row, column)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.MatrixStorage = NewData
    
    Set Add = Result
    
End Function

Public Function columns() As Long
    columns = this.MatrixData.columns
End Function

Public Property Let MatrixStorage(ByVal MatrixData As IMatrixStorage)
    Set this.MatrixData = MatrixData
End Property

Public Property Get MatrixStorage() As IMatrixStorage
    Set MatrixStorage = this.MatrixData
End Property

Public Property Let Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    this.MatrixData.Element(rowIndex, columnIndex) = Value
End Property

Public Property Get Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    Element = this.MatrixData.Element(rowIndex, columnIndex)
End Property

Public Function Equals(ByVal otherMatrix As IMatrix) As Boolean
    
    If Not Me.rows = otherMatrix.rows Or Not Me.columns = otherMatrix.columns Then
        Equals = False
    End If
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.rows - 1
        For column = 0 To this.MatrixData.columns - 1
            If Not Me.Element(row, column) = otherMatrix.Element(row, column) Then
                Equals = False
            End If
        Next column
    Next row
    
    Equals = True
    
End Function

Public Function IsSquare() As Boolean
    IsSquare = (this.MatrixData.rows = this.MatrixData.columns)
End Function

Public Function Multiply(ByVal otherMatrix As IMatrix) As DenseMatrix
    
    If Not this.MatrixData.columns = otherMatrix.rows Then MatrixErrors.ThrowMultiplicationError CLASS_NAME
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.rows, this.MatrixData.columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    Dim InnerProduct As Long
    
    For row = 0 To this.MatrixData.rows - 1
        For column = 0 To otherMatrix.columns - 1
            For InnerProduct = 0 To this.MatrixData.columns - 1
                NewData.Element(row, column) = NewData.Element(row, column) + _
                    this.MatrixData.Element(row, InnerProduct) * otherMatrix.Element(InnerProduct, column)
            Next InnerProduct
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.MatrixStorage = NewData
    
    Set Multiply = Result
    
End Function

Public Function ScalarMultiply(ByVal ScalarValue As Double) As DenseMatrix

    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.rows, this.MatrixData.columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.rows - 1
        For column = 0 To this.MatrixData.columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) * ScalarValue
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.MatrixStorage = NewData
    
    Set ScalarMultiply = Result

End Function

Public Function Subtract(ByVal otherMatrix As IMatrix) As DenseMatrix

    If Not this.MatrixData.rows = otherMatrix.rows Or Not this.MatrixData.columns = otherMatrix.columns Then
        MatrixErrors.ThrowSubtractionError CLASS_NAME
    End If
    
    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.rows, this.MatrixData.columns)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    For row = 0 To this.MatrixData.rows - 1
        For column = 0 To this.MatrixData.columns - 1
            NewData.Element(row, column) = this.MatrixData.Element(row, column) _
                - otherMatrix.Element(row, column)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.MatrixStorage = NewData
    
    Set Subtract = Result
    
End Function

Public Function rows() As Long
    rows = this.MatrixData.rows
End Function

Public Sub SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)

    Dim Temp As Double
    Dim column As Long
    
    For column = 0 To this.MatrixData.columns - 1
        With this.MatrixData
            Temp = .Element(secondRowIndex, column)
            .Element(secondRowIndex, column) = .Element(firstRowIndex, column)
            .Element(firstRowIndex, column) = Temp
        End With
    Next column

End Sub

Public Function Transpose() As DenseMatrix

    Dim factory As IMatrixStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.MatrixData)
    
    Dim NewData As IMatrixStorage
    Set NewData = factory.Create(this.MatrixData.columns, this.MatrixData.rows)
    
    Set factory = Nothing
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To this.MatrixData.columns - 1
        For column = 0 To this.MatrixData.rows - 1
            NewData.Element(row, column) = this.MatrixData.Element(column, row)
        Next column
    Next row
    
    Dim Result As IMatrix
    Set Result = New DenseMatrix
    
    Result.MatrixStorage = NewData
    
    Set Transpose = Result

End Function

' IMatrix Implementation

'@Description("Returns the sum of two matrices.")
Private Function IMatrix_Add(ByVal otherMatrix As IMatrix) As IMatrix
    Set IMatrix_Add = Me.Add(otherMatrix)
End Function

'@Description("Returns the number of columns in matrix.")
Private Function IMatrix_Columns() As Long
    IMatrix_Columns = Me.columns
End Function

Private Function IMatrix_Equals(ByVal otherMatrix As IMatrix) As Boolean
    IMatrix_Equals = Me.Equals(otherMatrix)
End Function

Private Property Let IMatrix_MatrixStorage(ByVal MatrixStorage As IMatrixStorage)
    Me.MatrixStorage = MatrixStorage
End Property

Private Property Get IMatrix_MatrixStorage() As IMatrixStorage
    IMatrix_MatrixStorage = Me.MatrixStorage
End Property

'@Description("Sets element value at specified location.")
Private Property Let IMatrix_Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    Me.Element(rowIndex, columnIndex) = Value
End Property

'@Description("Returns element value at specified location.")
Private Property Get IMatrix_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrix_Element = Me.Element(rowIndex, columnIndex)
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
    IMatrix_Rows = Me.rows
End Function

'@Description("Returns the product of scalar and matrix.")
Private Function IMatrix_ScalarMultiply(ByVal ScalarValue As Double) As IMatrix
    Set IMatrix_ScalarMultiply = Me.ScalarMultiply(ScalarValue)
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

