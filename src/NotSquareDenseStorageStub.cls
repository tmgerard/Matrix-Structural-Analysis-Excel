VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotSquareDenseStorageStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.LinearAlgebra.MatrixStorageStubs")
'@IgnoreModule
Option Explicit
Implements IMatrixStorage

Private matrixStub As IMatrixStorage

Const ROW_COUNT As Long = 2
Const COLUMN_COUNT As Long = 3
Private storageStub() As Double

Private Sub Class_Initialize()
    ReDim storageStub(0 To ROW_COUNT - 1, 0 To COLUMN_COUNT - 1)
    storageStub(0, 0) = 1
    storageStub(0, 1) = 1
    storageStub(0, 2) = 1
    storageStub(1, 0) = 1
    storageStub(1, 1) = 1
    storageStub(1, 2) = 1
End Sub

Private Sub Class_Terminate()
    Erase storageStub
End Sub

Private Sub IMatrixStorage_Clear()
    ReDim storageStub(0 To ROW_COUNT - 1, 0 To COLUMN_COUNT - 1)
End Sub

Private Property Get IMatrixStorage_Columns() As Long
    IMatrixStorage_Columns = COLUMN_COUNT
End Property

Private Property Let IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Double)
    storageStub(rowIndex, columnIndex) = value
End Property

Private Property Get IMatrixStorage_Element(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    IMatrixStorage_Element = storageStub(rowIndex, columnIndex)
End Property

Private Function IMatrixStorage_IsDense() As Boolean
    IMatrixStorage_IsDense = True
End Function

Private Function IMatrixStorage_IsMutable(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    IMatrixStorage_IsMutable = True
End Function

Private Property Get IMatrixStorage_Rows() As Long
    IMatrixStorage_Rows = ROW_COUNT
End Property

Private Sub IMatrixStorage_SetSize(ByVal Rows As Long, ByVal Columns As Long)

End Sub
