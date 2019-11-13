VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Vector")
Option Explicit
'Implements IMatrix
Implements IVector

Private Type TDenseVector
    VectorData As IVectorStorage
    operator As VectorOperator
End Type

Private this As TDenseVector
Private Const CLASS_NAME As String = "DenseVector"

Private Sub Class_Initialize()
    With this
        Set .VectorData = New IVectorStorage
        Set .operator = New VectorOperator
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .VectorData = Nothing
    End With
End Sub

Public Property Get Element(ByVal index As Long) As Double
    Element = this.VectorData.Element(index)
End Property

Public Property Let Element(ByVal index As Long, ByVal value As Double)
    this.VectorData.Element(index) = value
End Property

Public Property Get Length() As Long
    Length = this.VectorData.Length
End Property

Public Property Get Storage() As IVectorStorage
    Set Storage = this.VectorData
End Property

Public Property Let Storage(ByVal VectorData As IVectorStorage)
    Dim storageClone As ICloneable
    Set storageClone = VectorData
    Set this.VectorData = storageClone.Clone
End Property

Public Function Add(ByVal otherVector As IVector) As DenseVector
    Set Add = this.operator.Add(Me, otherVector)
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

Public Function ScalarMultiply(ByVal scalarValue As Double) As DenseVector
    ScalarMultiply = this.operator.ScalarMultiply(Me, scalarValue)
End Function

Public Function Subtract(ByVal otherVector As IVector) As DenseVector
    Set Subtract = this.operator.Subtract(Me, otherVector)
End Function

'IMatrix Implementation

' TODO: Implement IMatrix interface

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

Private Property Let IVector_Element(ByVal index As Long, ByVal value As Double)
    Me.Element(index) = value
End Property

Private Property Get IVector_Element(ByVal index As Long) As Double
    IVector_Element = Me.Element(index)
End Property

Private Function IVector_EuclideanDistance() As Double
    IVector_EuclideanDistance = Me.EuclideanDistance
End Function

Private Property Get IVector_Length() As Long
    IVector_Length = Me.Length
End Property

Private Property Get IVector_Orientation() As VectorOrientation

End Property

Private Function IVector_ScalarMultiply(ByVal scalarValue As Double) As IVector
    Set IVector_ScalarMultiply = Me.ScalarMultiply(scalarValue)
End Function

Private Property Get IVector_Storage() As IVectorStorage
    Set IVector_Storage = Me.Storage
End Property

Private Property Let IVector_Storage(ByVal VectorData As IVectorStorage)
    Me.Storage = VectorData
End Property

Private Function IVector_Subtract(ByVal otherVector As IVector) As IVector
    Set IVector_Subtract = Me.Subtract(otherVector)
End Function
