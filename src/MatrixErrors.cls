VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("StructuralAnalysis.Model.LinearAlgebra.Matrix.Errors")

Option Explicit

Public Enum MatrixError
    RowRange = vbObjectError + 100
    ColumnRange
    Addition
    Subtraction
    Multiplication
    RowSize
    ColumnSize
    StorageScheme
End Enum

Public Sub ThrowAdditionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Matrices not compatible for addition."
    
    Err.Raise Number:=MatrixError.Addition, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowColumnRangeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "The column index specified is outside the range of the matrix."
    
    Err.Raise Number:=MatrixError.ColumnRange, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowColumnSizeError(Optional ByVal ErrorSource As String = vbNullString)
    
    Const ErrorMessage As String = "Invalid number of columns."
    
    Err.Raise Number:=MatrixError.ColumnSize, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowRowRangeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "The row index specified is outside the range of the matrix."
    
    Err.Raise Number:=MatrixError.RowRange, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowRowSizeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Invalid number of rows."
    
    Err.Raise Number:=MatrixError.RowSize, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowMultiplicationError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Matrices not compatible for multiplication."
    
    Err.Raise Number:=MatrixError.Multiplication, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowStorageSchemeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Unsupported matrix storage scheme."
    
    Err.Raise Number:=MatrixError.StorageScheme, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowSubtractionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Matrices not compatible for subtraction."
    
    Err.Raise Number:=MatrixError.Subtraction, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub









