VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatrixStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Interface")
'@Interface IMatrixStorageFactory
Option Explicit

Public Enum MatrixStorageScheme
    DenseColumnMajor
    DenseRowMajor
    'Skyline
End Enum

'@Description("Create empty matrix.")
Public Function Create(ByVal rows As Long, ByVal columns As Long) As IMatrixStorage
End Function

'@Description("Create matrix with values from worksheet range.")
Public Function CreateFromRange(ByVal WorksheetRange As Range) As IMatrixStorage
End Function

'@Description("Create identity matrix")
Public Function CreateIdentity(ByVal rows As Long) As IMatrixStorage
End Function

'@Description("Create empty square matrix.")
Public Function CreateSquare(ByVal rows As Long) As IMatrixStorage
End Function
