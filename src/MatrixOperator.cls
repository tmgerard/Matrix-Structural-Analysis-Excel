VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Operations")

Private Const CLASS_NAME As String = "MatrixOperator"

Public Function Add(ByVal leftMatrix As IMatrix, ByVal rightMatrix As IMatrix) As IMatrix

    If Not leftMatrix.Rows = rightMatrix.Rows Or Not leftMatrix.Columns = rightMatrix.Columns Then
        MatrixErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim newMatrix As ICloneable
    Set newMatrix = leftMatrix
    
    Dim result As IMatrix
    Set result = newMatrix.Clone
    
    Dim row As Long
    Dim column As Long
    For row = 0 To leftMatrix.Rows - 1
        For column = 0 To leftMatrix.Columns - 1
            result.Element(row, column) = leftMatrix.Element(row, column) _
                + rightMatrix.Element(row, column)
        Next column
    Next row
    
    Set Add = result
    
End Function

Public Function Equals(ByVal leftMatrix As IMatrix, ByVal rightMatrix As IMatrix) As Boolean
    
    Dim result As Boolean
    result = True
    
    If Not leftMatrix Is rightMatrix Then
        result = False
        Exit Function
    End If
    
    If Not leftMatrix.Rows = rightMatrix.Rows Or Not leftMatrix.Columns = rightMatrix.Columns Then
        result = False
        Exit Function
    End If
    
    Dim row As Long
    Dim column As Long
    For row = 0 To leftMatrix.Rows - 1
        For column = 0 To leftMatrix.Columns - 1
            If Not leftMatrix.Element(row, column) = rightMatrix.Element(row, column) Then
                result = False
                Exit For
            End If
        Next column
    Next row
    
    Equals = result
    
End Function

Public Function Multiply(ByVal leftMatrix As IMatrix, ByVal rightMatrix As IMatrix) As IMatrix
    
    If Not leftMatrix.Columns = rightMatrix.Rows Then MatrixErrors.ThrowMultiplicationError CLASS_NAME
    
    Dim newMatrix As ICloneable
    Set newMatrix = leftMatrix
    
    Dim result As IMatrix
    Set result = newMatrix.Clone
    
    result.SetSize leftMatrix.Rows, rightMatrix.Columns
    
    Dim row As Long
    Dim column As Long
    Dim InnerProduct As Long
    
    For row = 0 To leftMatrix.Rows - 1
        For column = 0 To rightMatrix.Columns - 1
            For InnerProduct = 0 To leftMatrix.Columns - 1
                result.Element(row, column) = result.Element(row, column) + _
                    leftMatrix.Element(row, InnerProduct) * rightMatrix.Element(InnerProduct, column)
            Next InnerProduct
        Next column
    Next row
    
    Set Multiply = result
    
End Function

Public Function ScalarMultiply(ByVal leftMatrix As IMatrix, ByVal scalarValue As Double) As IMatrix
    
    Dim newMatrix As ICloneable
    Set newMatrix = leftMatrix
    
    Dim result As IMatrix
    Set result = newMatrix.Clone
    
    Dim row As Long
    Dim column As Long
    For row = 0 To leftMatrix.Rows - 1
        For column = 0 To leftMatrix.Columns - 1
            result.Element(row, column) = leftMatrix.Element(row, column) * scalarValue
        Next column
    Next row
    
    Set ScalarMultiply = result

End Function

Public Sub SwapRows(ByRef Matrix As IMatrix, ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)

    Dim Temp As Double
    Dim column As Long
    
    For column = 0 To Matrix.Columns - 1
        With Matrix
            Temp = .Element(secondRowIndex, column)
            .Element(secondRowIndex, column) = .Element(firstRowIndex, column)
            .Element(firstRowIndex, column) = Temp
        End With
    Next column

End Sub

Public Function Subtract(ByVal leftMatrix As IMatrix, ByVal rightMatrix As IMatrix) As IMatrix

    If Not leftMatrix.Rows = rightMatrix.Rows Or Not leftMatrix.Columns = rightMatrix.Columns Then
        MatrixErrors.ThrowSubtractionError CLASS_NAME
    End If
    
    Dim newMatrix As ICloneable
    Set newMatrix = leftMatrix
    
    Dim result As IMatrix
    Set result = newMatrix.Clone
    
    Dim row As Long
    Dim column As Long
    For row = 0 To leftMatrix.Rows - 1
        For column = 0 To leftMatrix.Columns - 1
            result.Element(row, column) = leftMatrix.Element(row, column) _
                - rightMatrix.Element(row, column)
        Next column
    Next row
    
    Set Subtract = result
    
End Function

Public Function Transpose(ByVal Matrix As IMatrix) As IMatrix

    Dim newMatrix As ICloneable
    Set newMatrix = Matrix
    
    Dim result As IMatrix
    Set result = newMatrix.Clone
    
    result.SetSize Matrix.Columns, Matrix.Rows
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To Matrix.Columns - 1
        For column = 0 To Matrix.Rows - 1
            result.Element(row, column) = Matrix.Element(column, row)
        Next column
    Next row
    
    Set Transpose = result

End Function
