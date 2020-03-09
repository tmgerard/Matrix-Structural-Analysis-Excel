VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Vector")
'@IgnoreModule ObjectVariableNotSet
Option Explicit
Implements ICloneable
Implements IMatrix
Implements IVector

Private Type TDenseVector
    vectorData As IVectorStorage
    operator As VectorOperator
End Type

Private this As TDenseVector
Private Const CLASS_NAME As String = "DenseVector"

Private Sub Class_Initialize()
    With this
        Set .vectorData = New IVectorStorage
        Set .operator = New VectorOperator
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .vectorData = Nothing
    End With
End Sub

Public Property Get Element(ByVal index As Long) As Double
    Element = this.vectorData.Element(index)
End Property

Public Property Let Element(ByVal index As Long, ByVal value As Double)
    this.vectorData.Element(index) = value
End Property

Public Property Get Length() As Long
    Length = this.vectorData.Length
End Property

Public Property Get Orientation() As VectorOrientation
    Orientation = this.vectorData.Orientation
End Property

Public Property Get Storage() As IVectorStorage
    Set Storage = this.vectorData
End Property

Public Property Set Storage(ByVal vectorData As IVectorStorage)
    Dim storageClone As ICloneable
    Set storageClone = vectorData
    Set this.vectorData = storageClone.Clone
End Property

Public Function Add(ByVal otherVector As IVector) As DenseVector
    Set Add = this.operator.Add(Me, otherVector)
End Function

Public Function Clone() As DenseVector
    Dim vectorData As ICloneable
    Set vectorData = this.vectorData
    
    Dim cloneVector As DenseVector
    Set cloneVector = New DenseVector
    
    Set cloneVector.Storage = vectorData.Clone
    
    Set Clone = cloneVector
End Function

Public Function CrossProduct(ByVal otherVector As IVector) As IVector
    Set CrossProduct = this.operator.CrossProduct(Me, otherVector)
End Function

Public Function DotProduct(ByVal otherVector As IVector) As Double
    DotProduct = this.operator.DotProduct(Me, otherVector)
End Function

Public Function EuclideanDistance() As Double
    EuclideanDistance = this.operator.EuclideanDistance(Me)
End Function

Public Function Equals(ByVal otherVector As IVector) As Boolean
    Equals = this.operator.Equals(Me, otherVector)
End Function

Public Function ScalarMultiply(ByVal scalarValue As Double) As DenseVector
    ScalarMultiply = this.operator.ScalarMultiply(Me, scalarValue)
End Function

'@Ignore UseMeaningfulName
Public Sub SwapValues(ByVal index1 As Long, ByVal index2 As Long)

    Dim temp As Double
    
    With this.vectorData
        temp = .Element(index1)
        .Element(index1) = .Element(index2)
        .Element(index2) = temp
    End With
    
End Sub

Public Function Subtract(ByVal otherVector As IVector) As DenseVector
    Set Subtract = this.operator.Subtract(Me, otherVector)
End Function

' ICloneable Implementation

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Me.Clone
End Function

' IMatrix Implementation

Private Function IMatrix_Add(ByVal otherMatrix As IMatrix) As IMatrix
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    Set IMatrix_Add = operator(Me, otherMatrix)
End Function

Private Function IMatrix_Columns() As Long
    If Me.Orientation = ColumnVector Then
        IMatrix_Columns = Me.Length
    Else
        IMatrix_Columns = 1
    End If
End Function

Private Function IMatrix_Equals(ByVal otherMatrix As IMatrix) As Boolean
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    IMatrix_Equals = operator.Equals(Me, otherMatrix)
End Function

Private Function IMatrix_IsSquare() As Boolean
    IMatrix_IsSquare = False
End Function

Private Function IMatrix_Multiply(ByVal otherMatrix As IMatrix) As IMatrix
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    Set IMatrix_Multiply = operator.Multiply(Me, otherMatrix)
End Function

Private Function IMatrix_Rows() As Long
    If Me.Orientation = RowVector Then
        IMatrix_Rows = Me.Length
    Else
        IMatrix_Rows = 1
    End If
End Function

Private Function IMatrix_ScalarMultiply(ByVal scalarValue As Double) As IMatrix
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    Set IMatrix_ScalarMultiply = operator.ScalarMultiply(Me, scalarValue)
End Function

Private Sub IMatrix_SetSize(ByVal Rows As Long, ByVal Columns As Long)
    If Rows > 1 And Columns > 1 Then VectorErrors.ThrowBadLengthError CLASS_NAME
    
    If Me.Orientation = ColumnVector Then
        this.vectorData.Length = Rows
    Else
        this.vectorData.Length = Columns
    End If
End Sub

Private Property Set IMatrix_storage(ByVal matrixData As IMatrixStorage)
    Set Me.Storage = matrixData
End Property

Private Property Get IMatrix_storage() As IMatrixStorage
    Set IMatrix_storage = Me.Storage
End Property

Private Function IMatrix_Subtract(ByVal otherMatrix As IMatrix) As IMatrix
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    Set IMatrix_Subtract = operator.Subtract(Me, otherMatrix)
End Function

Private Sub IMatrix_SwapRows(ByVal firstRowIndex As Long, ByVal secondRowIndex As Long)
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    operator.SwapRows Me, firstRowIndex, secondRowIndex
End Sub

Private Function IMatrix_Transpose() As IMatrix
    Dim operator As MatrixOperator
    Set operator = New MatrixOperator
    Set IMatrix_Transpose = operator.Transpose(Me)
End Function

' IVector Implementation

Private Function IVector_Add(ByVal otherVector As IVector) As IVector
    Set IVector_Add = Me.Add(otherVector)
End Function

Private Function IVector_CrossProduct(ByVal otherVector As IVector) As IVector
    Set IVector_CrossProduct = Me.CrossProduct(otherVector)
End Function

Private Function IVector_DotProduct(ByVal otherVector As IVector) As Double
    IVector_DotProduct = Me.DotProduct(otherVector)
End Function

'Private Property Let IVector_Element(ByVal index As Long, ByVal value As Double)
'    Me.Element(index) = value
'End Property
'
'Private Property Get IVector_Element(ByVal index As Long) As Double
'    IVector_Element = Me.Element(index)
'End Property

Private Function IVector_EuclideanDistance() As Double
    IVector_EuclideanDistance = Me.EuclideanDistance
End Function

Private Property Get IVector_Length() As Long
    IVector_Length = Me.Length
End Property

Private Property Get IVector_Orientation() As VectorOrientation
    IVector_Orientation = Me.Orientation
End Property

Private Function IVector_ScalarMultiply(ByVal scalarValue As Double) As IVector
    Set IVector_ScalarMultiply = Me.ScalarMultiply(scalarValue)
End Function

Private Property Get IVector_Storage() As IVectorStorage
    Set IVector_Storage = Me.Storage
End Property

Private Property Set IVector_Storage(ByVal vectorData As IVectorStorage)
    Set Me.Storage = vectorData
End Property

Private Function IVector_Subtract(ByVal otherVector As IVector) As IVector
    Set IVector_Subtract = Me.Subtract(otherVector)
End Function

'@Ignore UseMeaningfulName
Private Sub IVector_SwapValues(ByVal index1 As Long, ByVal index2 As Long)
    Me.SwapValues index1, index2
End Sub
