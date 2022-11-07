VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Equations")
Option Explicit

Private Type TVector
    Data() As Double
    Length As Long
End Type

Private this As TVector

Public Enum VectorErrors
    LengthMismatch = 500 + vbObjectError
End Enum

Public Function SetLength(ByRef vectorLength As Long) As Vector
    With this
        .Length = vectorLength
        ReDim .Data(vectorLength - 1)
    End With
    
    Set SetLength = Me
End Function

Public Property Get Length() As Long
    Length = this.Length
End Property

Public Property Get ValueAt(ByRef index As Long) As Double
    ValueAt = this.Data(index)
End Property

Public Property Let ValueAt(ByRef index As Long, ByRef value As Double)
    this.Data(index) = value
End Property

Public Function AddTo(ByRef index As Long, ByRef amount As Double) As Vector
    this.Data(index) = this.Data(index) + amount
    Set AddTo = Me
End Function

Public Function SetData(ByRef vecData() As Double) As Vector
    If Not ArrayLength(this.Data) = ArrayLength(vecData) Then
        Err.Raise Number:=VectorErrors.LengthMismatch, _
                  source:="Vector.SetData()", _
                  Description:="Vector data array length mismatch"
    End If
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = vecData(index)
    Next index
    
    Set SetData = Me
End Function

Public Function Equals(ByRef other As Vector) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Vector Then ' other is not a Vector
        Equals = False
    Else
        Dim index As Long
        For index = LBound(this.Data) To UBound(this.Data)
            If Not Doubles.Equal(Me.ValueAt(index), other.ValueAt(index)) Then
                Equals = False
                Exit Function
            End If
        Next index
        Equals = True
    End If

End Function

Public Function ScaleMatrix(ByRef factor As Double) As Matrix
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = this.Data(index) * factor
    Next index
    
    Set ScaleMatrix = Me
End Function

