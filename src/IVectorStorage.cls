VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVectorStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Vector.Interface")
Option Explicit

Public Enum VectorOrientation
    RowVector
    ColumnVector
End Enum

'@Description("Returns number of elements in vector.")
Public Property Get length() As Long
End Property

'@Description("Set the length of the vector")
Public Property Let length(ByVal length As Long)
End Property

Public Property Get Orientation() As VectorOrientation
End Property

'@Description("Returns element value at specified location.")
Public Property Get Element(ByVal index As Long) As Double
End Property

'@Description("Set element value at specified location.")
Public Property Let Element(ByVal index As Long, ByVal value As Double)
End Property

'@Description("Clear all vector elements.")
Public Sub Clear()
End Sub
