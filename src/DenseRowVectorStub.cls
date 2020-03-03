VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseRowVectorStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.LinearAlgebra.VectorStubs")
'@IgnoreModule
Option Explicit
Implements ICloneable
Implements IVector

Private vectorStub As IVector
Const ROW_COUNT As Long = 1
Const COLUMN_COUNT As Long = 3

Private Sub Class_Initialize()
    Set vectorStub = New DenseVector
    
    Dim Storage As IVectorStorage
    Set Storage = New DenseRowVectorStorage
    
    Storage.length = COLUMN_COUNT
    
    With vectorStub
        Set .Storage = Storage
        .Storage.Element(0) = 1
        .Storage.Element(1) = 1
        .Storage.Element(2) = 1
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

Private Function IVector_EuclideanDistance() As Double

End Function

Private Property Get IVector_Length() As Long
    IVector_Length = COLUMN_COUNT
End Property

Private Property Get IVector_Orientation() As VectorOrientation
    IVector_Orientation = RowVector
End Property

Private Function IVector_ScalarMultiply(ByVal scalarValue As Double) As IVector

End Function

Private Property Set IVector_Storage(ByVal RHS As IVectorStorage)
    Set IVector_Storage = RHS
End Property

Private Property Get IVector_Storage() As IVectorStorage
    Set IVector_Storage = vectorStub.Storage
End Property

Private Function IVector_Subtract(ByVal otherVector As IVector) As IVector

End Function

Private Sub IVector_SwapValues(ByVal index1 As Long, ByVal index2 As Long)

End Sub
