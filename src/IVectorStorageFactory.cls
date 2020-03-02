VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVectorStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Vector.Interface")
Option Explicit

Public Enum VectorStorageScheme
    Dense
    'Sparse
End Enum

'@Description("Create empty matrix.")
Public Function Create(ByVal length As Long) As IVectorStorage
End Function

'@Description("Create matrix with values from worksheet range.")
Public Function CreateFromRange(ByVal WorksheetRange As Range) As IVectorStorage
End Function
