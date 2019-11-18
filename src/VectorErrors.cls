VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Vector.Errors")

Option Explicit

Public Enum VectorError
    BadRange = vbObjectError + 200
    Addition
    Subtraction
    BadLength
    StorageScheme
    DotProduct
    CrossProduct
    CreateFromRange
End Enum

Public Sub ThrowAdditionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Vectors not compatible for addition."
    
    Err.Raise Number:=VectorError.Addition, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowBadRangeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "The index specified is outside the range of the vector."
    
    Err.Raise Number:=VectorError.BadRange, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowBadLengthError(Optional ByVal ErrorSource As String = vbNullString)
    
    Const ErrorMessage As String = "Invalid vector length."
    
    Err.Raise Number:=VectorError.BadLength, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowStorageSchemeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Unsupported vector storage scheme."
    
    Err.Raise Number:=VectorError.StorageScheme, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowSubtractionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Vectors not compatible for subtraction."
    
    Err.Raise Number:=VectorError.Subtraction, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowDotProductError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Vectors not compatible for dot product."
    
    Err.Raise Number:=VectorError.DotProduct, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowCrossProductError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "The cross product is only defined for vectors in three dimensions."
    
    Err.Raise Number:=VectorError.CrossProduct, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowCreateFromRangeError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Vector column and row count both greater than one."
    
    Err.Raise Number:=VectorError.CreateFromRange, _
              Source:=ErrorSource, _
              Description:=ErrorMessage

End Sub
