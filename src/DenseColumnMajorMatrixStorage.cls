VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseColumnMajorMatrixStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Storage")
Option Explicit
Implements IMatrixStorage
Implements ICloneable

Private Type TDenseColumnMajorMatrixStorage
    rows As Long
    columns As Long
    Data() As Double
End Type

Private this As TDenseColumnMajorMatrixStorage
Private Const CLASS_NAME As String = "DenseColumnMajorMatrixStorage"

Private Sub Class_Initialize()
    With this
        .rows = 0
        .columns = 0
    End With
End Sub

Private Sub Class_Terminate()
    With this
        .rows = 0
        .columns = 0
        Erase .Data
    End With
End Sub

'@Description("Return the number of columns in matrix")
Public Property Get columns() As Long
    columns = this.columns
End Property

'@Description("Return element at a given index.")
Public Property Let Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal Value As Double)
    
    If rowIndex < 0 Or rowIndex > this.rows - 1 Then MatrixErrors.ThrowRowRangeError ErrorSource:=CLASS_NAME
    If columnIndex < 0 Or columnIndex > this.columns - 1 Then MatrixErrors.ThrowColumnRangeError ErrorSource:=CLASS_NAME
    
    this.Data(ElementIndex(rowIndex, columnIndex)) = Value
    
End Property

'@Description("Set element at a given index.")
Public Property Get Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double

    If rowIndex < 0 Or rowIndex > this.rows - 1 Then MatrixErrors.ThrowRowRangeError ErrorSource:=CLASS_NAME
    If columnIndex < 0 Or columnIndex > this.columns - 1 Then MatrixErrors.ThrowColumnRangeError ErrorSource:=CLASS_NAME
    
    Element = this.Data(ElementIndex(rowIndex, columnIndex))

End Property

'@Description("Return the number of rows in matrix")
Public Property Get rows() As Long
    rows = this.rows
End Property

'@Description("Set all matrix elements to zero.")
Public Sub Clear()
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = 0
    Next index

End Sub

'@Description("Return copy of matrix storage object")
Public Function Clone() As DenseColumnMajorMatrixStorage

    Dim NewMatrix As DenseColumnMajorMatrixStorage
    Set NewMatrix = New DenseColumnMajorMatrixStorage
    
    NewMatrix.SetSize Me.rows, Me.columns
    
    Dim rowIndex As Long
    Dim columnIndex As Long
    For rowIndex = 0 To this.rows - 1
        For columnIndex = 0 To this.columns - 1
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
Public Sub SetSize(ByVal rows As Long, ByVal columns As Long)

    If rows <= 0 Then MatrixErrors.ThrowRowSizeError
    If columns <= 0 Then MatrixErrors.ThrowColumnSizeError
    
    With this
        .rows = rows
        .columns = columns
        ReDim .Data(rows * columns - 1)
    End With

End Sub

'@Description("Retrieves element from data array based on the row and column index.")
Private Function ElementIndex(ByVal rowIndex As Long, ByVal columnIndex As Long) As Long
    ElementIndex = columnIndex * this.rows + rowIndex
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
    IMatrixStorage_Columns = Me.columns
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

Private Sub IMatrixStorage_SetSize(ByVal rows As Long, ByVal columns As Long)
    Me.SetSize rows:=rows, columns:=columns
End Sub

Private Property Get IMatrixStorage_Rows() As Long
    IMatrixStorage_Rows = Me.rows
End Property

