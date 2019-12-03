VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatrixStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.LinearAlgebra.Matrix.Interface")
'@Interface IMatrixStorage
Option Explicit

'@Description("Returns number of rows in matrix.")
Public Property Get Rows() As Long
End Property

'@Description("Returns number of columns in matrix.")
Public Property Get Columns() As Long
End Property

'@Description("Returns element value at specified location.")
Public Property Get Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
End Property

'@Description("Set element value at specified location.")
Public Property Let Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Double)
End Property

'@Description("Check element at a given index to check if it is mutable")
Public Function IsMutable(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
End Function

'@Description("Check that matrix is dense.")
Public Function IsDense() As Boolean
End Function

'@Description("Clear all matrix elements.")
Public Sub Clear()
End Sub

'@Description("Set the size of the matrix")
Public Sub SetSize(ByVal Rows As Long, ByVal Columns As Long)
End Sub
