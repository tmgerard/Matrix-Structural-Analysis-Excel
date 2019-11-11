VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Interface")
'@Interface IMatrix
Option Explicit

'@Description("Returns the sum of two matrices.")
Public Function Add(ByVal otherMatrix As IMatrix) As IMatrix
End Function

'@Description("Returns the number of columns in matrix.")
Public Function Columns() As Long
End Function

'@Description("Inject matrix data storage object.")
Public Property Let Storage(ByVal matrixData As IMatrixStorage)
End Property

Public Property Get Storage() As IMatrixStorage
End Property

'@Description("Returns element value at specified location.")
Public Property Get Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
End Property

'@Description("Set element value at specified location.")
Public Property Let Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Double)
End Property

'@Description("Check equality of two matrices")
Public Function Equals(ByVal otherMatrix As IMatrix) As Boolean
End Function

'@Description("Returns true if the matrix is square.")
Public Function IsSquare() As Boolean
End Function

'@Description("Returns the product of two matrices.")
Public Function Multiply(ByVal otherMatrix As IMatrix) As IMatrix
End Function

'@Description("Returns the number of rows in matrix.")
Public Function Rows() As Long
End Function

'@Description("Returns the product of scalar and matrix.")
Public Function ScalarMultiply(ByVal scalarValue As Double) As IMatrix
End Function

'@Description("Set the size of the matrix")
Public Sub SetSize(ByVal Rows As Long, ByVal Columns As Long)
End Sub

'@Description("Returns the difference of two matrices.")
Public Function Subtract(ByVal otherMatrix As IMatrix) As IMatrix
End Function

'@Description("Exchange the position of two rows within a matrix")
Public Sub SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)
End Sub

'@Description("Return the transposes of a matrix")
Public Function Transpose() As IMatrix
End Function
