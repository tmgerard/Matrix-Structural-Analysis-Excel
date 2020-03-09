VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Vector.Interface")
Option Explicit

Public Function Add(ByVal otherVector As IVector) As IVector
End Function

Public Function CrossProduct(ByVal otherVector As IVector) As IVector
End Function

Public Function DotProduct(ByVal otherVector As IVector) As Double
End Function

'Public Property Get Element(ByVal index As Long) As Double
'End Property
'
'Public Property Let Element(ByVal index As Long, ByVal value As Double)
'End Property

Public Function EuclideanDistance() As Double
End Function

Public Property Get Length() As Long
End Property

Public Property Get Orientation() As VectorOrientation
End Property

Public Function Subtract(ByVal otherVector As IVector) As IVector
End Function

Public Function ScalarMultiply(ByVal scalarValue As Double) As IVector
End Function

'@Ignore UseMeaningfulName
Public Sub SwapValues(ByVal index1 As Long, ByVal index2 As Long)
End Sub

Public Property Set Storage(ByVal vectorData As IVectorStorage)
End Property

Public Property Get Storage() As IVectorStorage
End Property

