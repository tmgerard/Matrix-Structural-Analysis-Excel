VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatrixOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Matrix.Interface")
'@Interface IMatrixOutput
Option Explicit

Private Const DEFAULT_NUMBER_FORMAT As String = "#0.0000"
Private Const DEFAULT_FILENAME As String = "matrix_output"

Public Sub ToImmediateWindow(ByVal matrixData As IMatrixStorage, Optional ByVal OutputFormat As String = DEFAULT_NUMBER_FORMAT)
End Sub

Public Sub ToFile(ByVal matrixData As IMatrixStorage, Optional ByVal FileName As String = DEFAULT_FILENAME, Optional ByVal OutputFormat As String = DEFAULT_NUMBER_FORMAT)
End Sub

Public Sub ToWorksheet(ByVal matrixData As IMatrixStorage, ByVal Cell As Range, Optional ByVal OutputFormat As String = DEFAULT_NUMBER_FORMAT)
End Sub
