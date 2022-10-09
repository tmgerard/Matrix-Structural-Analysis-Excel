VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rectangle2DSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TRectangle2DSize
    Width As Double
    Height As Double
End Type

Private this As TRectangle2DSize

Public Property Get Height() As Double
    Height = this.Height
End Property

Public Property Let Height(ByRef value As Double)
    this.Height = value
End Property

Public Property Get Width() As Double
    Width = this.Width
End Property

Public Property Let Width(ByRef value As Double)
    this.Width = value
End Property

Public Function Equals(ByRef other As Rectangle2DSize) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Rectangle2DSize Then ' other is not a Rectangle2DSize
        Equals = False
    Else
        Equals = Doubles.Equal(this.Width, other.Width) And Doubles.Equal(this.Height, other.Height)
    End If

End Function
