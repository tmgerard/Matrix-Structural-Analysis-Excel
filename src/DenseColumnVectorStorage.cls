VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseColumnVectorStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Matrix.Storage")
Option Explicit
Implements IMatrixStorage
Implements IVectorStorage
Implements ICloneable

Private Type TDenseColumnVectorStorage
    Rows As Long
    Columns As Long
    Length As Long
    Data() As Double
    Orientation As VectorOrientation
End Type

Private this As TDenseColumnVectorStorage
Private Const COLUMN_COUNT As Long = 1
Private Const CLASS_NAME As String = "DenseColumnVectorStorage"

Private Sub Class_Initialize()
    With this
        .Rows = 0
        .Columns = COLUMN_COUNT
        .Length = 0
        .Orientation = ColumnVector
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Erase .Data
    End With
End Sub

Public Property Get Columns() As Long
    Columns = this.Columns
End Property

Public Property Let Columns(ByVal Columns As Long)
    If Columns <= 0 Or Columns > COLUMN_COUNT Then MatrixErrors.ThrowColumnSizeError
End Property

'@Description("Return the number of elements in vector")
Public Property Get Length() As Long
    Length = this.Length
End Property

Public Property Get Rows() As Long
    Rows = this.Rows
End Property

'@Description("Set the number of elements to store in the vector")
Public Property Let Length(ByVal Length As Long)
    If Length <= 0 Then VectorErrors.ThrowBadLengthError
    
    With this
        .Length = Length
        .Rows = Length
        ReDim .Data(Length - 1)
    End With
End Property

Public Property Get Orientation() As VectorOrientation
    Orientation = this.Orientation
End Property

'@Description("Return element at a given index")
Public Property Get Element(ByVal index As Long) As Double
    If index < 0 Or index > this.Length - 1 Then VectorErrors.ThrowBadRangeError ErrorSource:=CLASS_NAME
    Element = this.Data(index)
End Property

'@Description("Set element value at a given index")
Public Property Let Element(ByVal index As Long, ByVal value As Double)
    If index < 0 Or index > this.Length - 1 Then VectorErrors.ThrowBadLengthError ErrorSource:=CLASS_NAME
    this.Data(index) = value
End Property

Public Function IsDense() As Boolean
    IsDense = True
End Function

Public Function IsMutable() As Boolean
    IsMutable = True
End Function

'@Description("Create a copy of vector object.")
Public Function Clone() As Object

    Dim newVector As DenseColumnVectorStorage
    Set newVector = New DenseColumnVectorStorage
    
    newVector.Length = Me.Length
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        newVector.Element(index) = this.Data(index)
    Next index
    
    Set Clone = newVector
    
End Function

'@Description("Set all vector elements to zero.")
Public Sub Clear()
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = 0#
    Next index
End Sub

' ICloneable Implementation

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Me.Clone
End Function

'IMatrixStorage Implementation

Private Sub IMatrixStorage_Clear()
    Me.Clear
End Sub

Private Property Get IMatrixStorage_Columns() As Long
    IMatrixStorage_Columns = Me.Columns
End Property

'@Ignore ParameterNotUsed
Private Property Let IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Double)
    Me.Element(rowIndex) = value
End Property

'@Ignore ParameterNotUsed
Private Property Get IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrixStorage_Element = Me.Element(rowIndex)
End Property

Private Function IMatrixStorage_IsDense() As Boolean
    IMatrixStorage_IsDense = Me.IsDense
End Function

'@Ignore ParameterNotUsed
Private Function IMatrixStorage_IsMutable(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    IMatrixStorage_IsMutable = Me.IsMutable()
End Function

Private Property Get IMatrixStorage_Rows() As Long
    IMatrixStorage_Rows = Me.Rows
End Property

'@Ignore ParameterNotUsed
Private Sub IMatrixStorage_SetSize(ByVal Rows As Long, ByVal Columns As Long)
    Me.Length = Rows
End Sub

' IVectorStorage Implementation

Private Sub IVectorStorage_Clear()
    Me.Clear
End Sub

Private Property Let IVectorStorage_Element(ByVal index As Long, ByVal value As Double)
    Me.Element(index) = value
End Property

Private Property Get IVectorStorage_Element(ByVal index As Long) As Double
    IVectorStorage_Element = Me.Element(index)
End Property

Private Property Get IVectorStorage_Length() As Long
    IVectorStorage_Length = Me.Length
End Property

Private Property Let IVectorStorage_Length(ByVal Length As Long)
    Me.Length = Length
End Property

Private Property Get IVectorStorage_Orientation() As VectorOrientation
    IVectorStorage_Orientation = Me.Orientation
End Property
