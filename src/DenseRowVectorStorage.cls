VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseRowVectorStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Matrix.Storage")
Option Explicit
Implements IMatrixStorage
Implements IVectorStorage
Implements ICloneable

Private Type TDenseRowVectorStorage
    Rows As Long
    Columns As Long
    length As Long
    Data() As Double
    Orientation As VectorOrientation
End Type

Private this As TDenseRowVectorStorage
Private Const ROW_COUNT As Long = 1
Private Const CLASS_NAME As String = "DenseRowVectorStorage"

Private Sub Class_Initialize()
    With this
        .Rows = ROW_COUNT
        .Columns = 0
        .length = 0
        .Orientation = RowVector
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
    If Columns <= 0 Then MatrixErrors.ThrowColumnSizeError
    
    With this
        .length = Columns
        .Columns = Columns
        ReDim .Data(Columns - 1)
    End With
End Property

'@Description("Return the number of elements in vector")
Public Property Get length() As Long
    length = this.length
End Property

Public Property Get Rows() As Long
    Rows = ROW_COUNT
End Property

'@Description("Set the number of elements to store in the vector")
Public Property Let length(ByVal length As Long)
    If length <= 0 Then VectorErrors.ThrowBadLengthError
    
    With this
        .length = length
        .Columns = length
        ReDim .Data(length - 1)
    End With
End Property

Public Property Get Orientation() As VectorOrientation
    Orientation = this.Orientation
End Property

'@Description("Return element at a given index")
Public Property Get Element(ByVal index As Long) As Double
    If index < 0 Or index > this.length - 1 Then VectorErrors.ThrowBadRangeError ErrorSource:=CLASS_NAME
    Element = this.Data(index)
End Property

'@Description("Set element value at a given index")
Public Property Let Element(ByVal index As Long, ByVal value As Double)
    If index < 0 Or index > this.length - 1 Then VectorErrors.ThrowBadLengthError ErrorSource:=CLASS_NAME
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

    Dim newVector As DenseRowVectorStorage
    Set newVector = New DenseRowVectorStorage
    
    newVector.length = Me.length
    
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
    Me.Element(columnIndex) = value
End Property

'@Ignore ParameterNotUsed
Private Property Get IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrixStorage_Element = Me.Element(columnIndex)
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
    Me.length = Columns
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
    IVectorStorage_Length = Me.length
End Property

Private Property Let IVectorStorage_Length(ByVal length As Long)
    Me.length = length
End Property

Private Property Get IVectorStorage_Orientation() As VectorOrientation
    IVectorStorage_Orientation = Me.Orientation
End Property
