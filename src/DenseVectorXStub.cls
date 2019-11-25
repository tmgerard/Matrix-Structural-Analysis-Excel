VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseVectorXStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Linear Algebra.Vector Stubs")
Option Explicit
Implements ICloneable
Implements IVector

Private vectorStub As IVector
Const ROW_COUNT As Long = 4
Const COLUMN_COUNT As Long = 1

Private Sub Class_Initialize()
    Set vectorStub = New DenseVector
    
    Dim Storage As IVectorStorage
    Set Storage = New DenseRowVectorStorage
    
    Storage.length = ROW_COUNT
    
    With vectorStub
        .Storage = Storage
        .Element(0) = 1
        .Element(1) = 1
        .Element(2) = 1
        .Element(3) = 1
    End With
End Sub

Private Sub Class_Terminate()
    Set vectorStub = Nothing
End Sub

Private Function ICloneable_Clone() As Object
    Dim cloned As ICloneable
    Set cloned = vectorStub
    Set ICloneable_Clone = cloned.Clone
End Function

Private Function IVector_Add(ByVal otherVector As IVector) As IVector

End Function

Private Function IVector_CrossProduct(ByVal otherVector As IVector) As IVector

End Function

Private Function IVector_DotProduct(ByVal otherVector As IVector) As Double

End Function

Private Property Let IVector_Element(ByVal index As Long, ByVal value As Double)
    vectorStub.Element(index) = value
End Property

Private Property Get IVector_Element(ByVal index As Long) As Double
    IVector_Element = vectorStub.Element(index)
End Property

Private Function IVector_EuclideanDistance() As Double

End Function

Private Property Get IVector_Length() As Long
    IVector_Length = ROW_COUNT
End Property

Private Property Get IVector_Orientation() As VectorOrientation
    IVector_Orientation = ColumnVector
End Property

Private Function IVector_ScalarMultiply(ByVal scalarValue As Double) As IVector

End Function

Private Property Let IVector_Storage(ByVal RHS As IVectorStorage)

End Property

Private Property Get IVector_Storage() As IVectorStorage

End Property

Private Function IVector_Subtract(ByVal otherVector As IVector) As IVector

End Function

Private Sub IVector_SwapValues(ByVal index1 As Long, ByVal index2 As Long)

End Sub
