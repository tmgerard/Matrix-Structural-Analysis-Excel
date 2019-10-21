VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVectorStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Matrix.Interface")
Option Explicit

'@Description("Returns number of elements in vector.")
Public Property Get Length() As Long
End Property

'@Description("Returns element value at specified location.")
Public Property Get Element(ByVal index As Long) As Double
End Property

'@Description("Set element value at specified location.")
Public Property Let Element(ByVal index As Long, ByVal Value As Double)
End Property

'@Description("Clear all vector elements.")
Public Sub Clear()
End Sub

'@Description("Set the length of the vector")
Public Sub SetLength(ByVal Length As Long)
End Sub
