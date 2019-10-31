VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseVectorStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Vector.Storage")
Option Explicit
Implements IVectorStorage
Implements ICloneable

Private Type TDenseVectorStorage
    Length As Long
    Data() As Double
End Type

Private this As TDenseVectorStorage
Private Const CLASS_NAME As String = "DenseVectorStorage"

Private Sub Class_Initialize()
    With this
        .Length = 0
    End With
End Sub

Private Sub Class_Terminate()
    With this
        .Length = 0
        Erase .Data
    End With
End Sub

'@Description("Return the number of elements in vector")
Public Property Get Length() As Long
    Length = this.Length
End Property

'@Description("Return element at a given index")
Public Property Get Element(ByVal index As Long) As Double
    If index < 0 Or index > this.Length - 1 Then VectorErrors.ThrowBadRangeError ErrorSource:=CLASS_NAME
    Element = this.Data(index)
End Property

'@Description("Set element value at a given index")
Public Property Let Element(ByVal index As Long, ByVal Value As Double)
    If index < 0 Or index > this.Length - 1 Then VectorErrors.ThrowBadLengthError ErrorSource:=CLASS_NAME
    this.Data(index) = Value
End Property

'@Description("Set the number of elements to store in the vector")
Public Sub SetLength(ByVal Length As Long)
    If Length <= 0 Then VectorErrors.ThrowBadLengthError
    
    With this
        .Length = Length
        ReDim .Data(Length - 1)
    End With
End Sub

'@Description("Create a copy of vector object.")
Public Function Clone() As Object

    Dim NewVector As DenseVectorStorage
    Set NewVector = New DenseVectorStorage
    
    NewVector.SetLength Me.Length
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        NewVector.Element(index) = this.Data(index)
    Next index
    
    Set Clone = NewVector
    
End Function


'@Description("Set all vector elements to zero.")
Public Sub Clear()
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = 0
    Next index
End Sub

' ICloneable Implementation

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Me.Clone
End Function

' IVectorStorage Implementation

Private Sub IVectorStorage_Clear()
    Me.Clear
End Sub

Private Property Let IVectorStorage_Element(ByVal index As Long, ByVal Value As Double)
    Me.Element(index) = Value
End Property

Private Property Get IVectorStorage_Element(ByVal index As Long) As Double
    IVectorStorage_Element = Me.Element(index)
End Property

Private Property Get IVectorStorage_Length() As Long
    IVectorStorage_Length = Me.Length
End Property

Private Sub IVectorStorage_SetLength(ByVal Length As Long)
    Me.SetLength Length:=Length
End Sub
