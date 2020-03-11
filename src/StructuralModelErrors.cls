VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructuralModelErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.Errors")

Option Explicit

Public Enum StructuralModelError
    BadElementLength = vbObjectError + 400
    BadElementLoadLocation
    UnstableStructure
End Enum

Public Sub ThrowBadElementLengthError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Element length must be a positive, non-zero value."
    
    Err.Raise Number:=StructuralModelError.BadElementLength, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

Public Sub ThrowBadLoadLocationError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Element load is not located on the element."
    
    Err.Raise Number:=StructuralModelError.BadElementLength, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
              
End Sub
