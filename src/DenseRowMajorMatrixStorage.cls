VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseRowMajorMatrixStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Matrix.Storage")
Option Explicit
Implements IMatrixStorage
Implements ICloneable

Private Type TDenseRowMajorMatrixStorage
    Rows As Long
    COLUMNS As Long
    Data() As Double
End Type

Private this As TDenseRowMajorMatrixStorage
Private Const CLASS_NAME As String = "DenseRowMajorMatrixStorage"

Private Sub Class_Initialize()
    With this
        .Rows = 0
        .COLUMNS = 0
    End With
End Sub

Private Sub Class_Terminate()
    With this
        .Rows = 0
        .COLUMNS = 0
        Erase .Data
    End With
End Sub

'@Description("Return the number of columns in matrix")
Public Property Get COLUMNS() As Long
    COLUMNS = this.COLUMNS
End Property

'@Description("Return element at a given index.")
Public Property Let Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    
    If rowIndex < 0 Or rowIndex > this.Rows - 1 Then MatrixErrors.ThrowRowRangeError ErrorSource:=CLASS_NAME
    If columnIndex < 0 Or columnIndex > this.COLUMNS - 1 Then MatrixErrors.ThrowColumnRangeError ErrorSource:=CLASS_NAME
    
    this.Data(ElementIndex(rowIndex, columnIndex)) = Value
    
End Property

'@Description("Set element at a given index.")
Public Property Get Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double

    If rowIndex < 0 Or rowIndex > this.Rows - 1 Then MatrixErrors.ThrowRowRangeError ErrorSource:=CLASS_NAME
    If columnIndex < 0 Or columnIndex > this.COLUMNS - 1 Then MatrixErrors.ThrowColumnRangeError ErrorSource:=CLASS_NAME
    
    Element = this.Data(ElementIndex(rowIndex, columnIndex))

End Property

'@Description("Return the number of rows in matrix")
Public Property Get Rows() As Long
    Rows = this.Rows
End Property

'@Description("Set all matrix elements to zero.")
Public Sub Clear()
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = 0
    Next index

End Sub

'@Description("Return copy of matrix storage object")
Public Function Clone() As DenseRowMajorMatrixStorage

    Dim NewMatrix As DenseRowMajorMatrixStorage
    Set NewMatrix = New DenseRowMajorMatrixStorage
    
    NewMatrix.SetSize Me.Rows, Me.COLUMNS
    
    Dim rowIndex As Long
    Dim columnIndex As Long
    For rowIndex = 0 To this.Rows - 1
        For columnIndex = 0 To this.COLUMNS - 1
            NewMatrix.Element(rowIndex, columnIndex) = Me.Element(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
    
    Set Clone = NewMatrix
    
End Function

'@Description("Check that value at a given index can be changed")
'@Ignore ParameterNotUsed
Public Function IsMutable(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    IsMutable = True
End Function

'@Description("Set the size of the matrix.")
Public Sub SetSize(ByVal Rows As Long, ByVal COLUMNS As Long)

    If Rows <= 0 Then MatrixErrors.ThrowRowSizeError
    If COLUMNS <= 0 Then MatrixErrors.ThrowColumnSizeError
    
    With this
        .Rows = Rows
        .COLUMNS = COLUMNS
        ReDim .Data(Rows * COLUMNS - 1)
    End With

End Sub

'@Description("Retrieves element from data array based on the row and column index.")
Private Function ElementIndex(ByVal rowIndex As Long, ByVal columnIndex As Long) As Long
    ElementIndex = rowIndex * this.COLUMNS + columnIndex
End Function

' ICloneable Implementation

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Me.Clone
End Function

' IMatrixStorage Implementation

Private Sub IMatrixStorage_Clear()
    Me.Clear
End Sub

Private Property Get IMatrixStorage_Columns() As Long
    IMatrixStorage_Columns = Me.COLUMNS
End Property

Private Property Let IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    Me.Element(rowIndex, columnIndex) = Value
End Property

Private Property Get IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrixStorage_Element = Me.Element(rowIndex, columnIndex)
End Property

Private Function IMatrixStorage_IsDense() As Boolean
    IMatrixStorage_IsDense = True
End Function

Private Function IMatrixStorage_IsMutable(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    IMatrixStorage_IsMutable = Me.IsMutable(rowIndex, columnIndex)
End Function

Private Sub IMatrixStorage_SetSize(ByVal Rows As Long, ByVal COLUMNS As Long)
    Me.SetSize Rows:=Rows, COLUMNS:=COLUMNS
End Sub

Private Property Get IMatrixStorage_Rows() As Long
    IMatrixStorage_Rows = Me.Rows
End Property
