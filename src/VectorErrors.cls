VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Model.Matrix.Errors")

Option Explicit

Public Enum VectorError
    BadRange = vbObjectError + 200
    Addition
    Subtraction
    BadLength
    StorageScheme
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
    
    Err.Raise Number:=VectorError.Addition, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub


