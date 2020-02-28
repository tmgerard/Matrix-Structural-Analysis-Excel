VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructuralCrossSectionErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.StructuralCrossSections.Errors")

Option Explicit

Public Enum CrossSectionError
    dimension = vbObjectError + 400
End Enum

Public Sub ThrowDimensionError(Optional ByVal ErrorSource As String = vbNullString)

    Const ErrorMessage As String = "Structural cross section dimension cannot be zero or negative."
    
    Err.Raise Number:=CrossSectionError.dimension, _
              Source:=ErrorSource, _
              Description:=ErrorMessage
End Sub

