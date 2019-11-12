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
Implements IVector

Private Type TDenseVector
    VectorData As IVectorStorage
End Type

Private this As TDenseVector
Private Const CLASS_NAME As String = "DenseVector"

Private Sub Class_Initialize()
    With this
        Set .VectorData = New IVectorStorage
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

    If Not this.VectorData.Length = otherVector.Length Then
        VectorErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.VectorData)
    
    Dim newData As IVectorStorage
    Set newData = factory.Create(this.VectorData.Length)
    
    Set factory = Nothing
    
    Dim index As Long
    For index = 0 To newData.Length - 1
        newData.Element(index) = this.VectorData.Element(index) + _
            otherVector.Element(index)
    Next index
    
    Dim result As IVector
    Set result = New DenseVector
    
    result.Storage = newData
    
    Set Add = result

End Function

Public Function Subtract(ByVal otherVector As IVector) As DenseVector

    If Not this.VectorData.Length = otherVector.Length Then
        VectorErrors.ThrowAdditionError CLASS_NAME
    End If
    
    Dim factory As IVectorStorageFactory
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(this.VectorData)
    
    Dim newData As IVectorStorage
    Set newData = factory.Create(this.VectorData.Length)
    
    Set factory = Nothing
    
    Dim index As Long
    For index = 0 To newData.Length - 1
        newData.Element(index) = this.VectorData.Element(index) - _
            otherVector.Element(index)
    Next index
    
    Dim result As IVector
    Set result = New DenseVector
    
    result.Storage = newData
    
    Set Subtract = result

End Function

' IVector Implementation

Private Function IVector_Add(ByVal otherVector As IVector) As IVector
    Set IVector_Add = Me.Add(otherVector)
End Function

Private Function IVector_DotProduct(ByVal otherVector As IVector) As Double
    IVector_DotProduct = 0
End Function

Private Property Let IVector_Element(ByVal index As Long, ByVal value As Double)
    Me.Element(index) = value
End Property

Private Property Get IVector_Element(ByVal index As Long) As Double
    IVector_Element = 0
End Property

Private Function IVector_EuclideanNorm() As Double
    IVector_EuclideanNorm = 0
End Function

Private Property Get IVector_Length() As Long
    IVector_Length = Me.Length
End Property

Private Function IVector_ScalarMultiply(ByVal scalarValue As Double) As IVector
    Set IVector_ScalarMultiply = Nothing
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
