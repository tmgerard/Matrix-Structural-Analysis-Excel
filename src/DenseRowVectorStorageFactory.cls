VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseRowVectorStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Factory")
Option Explicit
Implements IMatrixStorageFactory
Implements IVectorStorageFactory

Private Type TDenseRowVectorStorageFactory
    vectorData As DenseRowVectorStorage
End Type

Private this As TDenseRowVectorStorageFactory
Private Const CLASS_NAME As String = "DenseRowVectorStorageFactory"

Private Sub Class_Initialize()
    Set this.vectorData = New DenseRowVectorStorage
End Sub

Private Sub Class_Terminate()
    Set this.vectorData = Nothing
End Sub

'@Desctiption("Create an empty vector of a given length.")
Public Function Create(ByVal length As Long) As DenseRowVectorStorage
    this.vectorData.length = length
    Set Create = this.vectorData
End Function

'@Description("Create vector from a worksheet range")
Public Function CreateFromRange(ByVal WorksheetRange As Range) As DenseRowVectorStorage
    
    ' a single column or row must be selected
    If Not IsRangeAVector(WorksheetRange) Then
        VectorErrors.ThrowCreateFromRangeError CLASS_NAME
    End If
    
    If WorksheetRange.Rows.Count = 1 Then
        Set CreateFromRange = VectorFromRowRange(WorksheetRange)
    Else
        Set CreateFromRange = VectorFromColumnRange(WorksheetRange)
    End If

End Function

'@Description("Check that a single row or column is in the given range")
Private Function IsRangeAVector(ByVal WorksheetRange As Range) As Boolean
    
    Dim Rows As Long
    Rows = WorksheetRange.Rows.Count
    
    Dim Columns As Long
    Columns = WorksheetRange.Columns.Count
    
    If Rows > 1 And Columns > 1 Then
        IsRangeAVector = False
    Else
        IsRangeAVector = True
    End If
    
End Function

'@Description("Create vector from a row oriented worksheet range")
Private Function VectorFromRowRange(ByVal WorksheetRange As Range) As DenseRowVectorStorage
    
    Dim rangeRow As Long
    rangeRow = WorksheetRange.row
    
    Dim Columns As Long
    Columns = WorksheetRange.Columns.Count
    
    this.vectorData.length = Columns
    
    Dim column As Long
    For column = 0 To Columns - 1
        this.vectorData.Element(column) = CDbl(WorksheetRange.Cells(rangeRow, column + 1).Value2)
    Next column
    
    Set VectorFromRowRange = this.vectorData
    
End Function

'@Description("Create vector from a column oriented worksheet range")
Private Function VectorFromColumnRange(ByVal WorksheetRange As Range) As DenseRowVectorStorage

    Dim Rows As Long
    Rows = WorksheetRange.Rows.Count
    
    Dim rangeColumn As Long
    rangeColumn = WorksheetRange.column
    
    this.vectorData.length = Rows
    
    Dim row As Long
    For row = 0 To Rows - 1
        this.vectorData.Element(row) = CDbl(WorksheetRange.Cells(row + 1, rangeColumn).Value2)
    Next row
    
    Set VectorFromColumnRange = this.vectorData

End Function

' IMatrixStorageFactory Implementation

Private Function IMatrixStorageFactory_Create(ByVal Rows As Long, ByVal Columns As Long) As IMatrixStorage
    Set IMatrixStorageFactory_Create = Me.Create(Columns)
End Function

Private Function IMatrixStorageFactory_CreateFromRange(ByVal WorksheetRange As Range) As IMatrixStorage
    Set IMatrixStorageFactory_CreateFromRange = Me.CreateFromRange(WorksheetRange)
End Function

Private Function IMatrixStorageFactory_CreateIdentity(ByVal Rows As Long) As IMatrixStorage
    ' TODO: Throw error if trying to create identity vector
End Function

Private Function IMatrixStorageFactory_CreateSquare(ByVal Rows As Long) As IMatrixStorage
    ' TODO: Throw error if trying to create square vector
End Function

' IVectorStorageFactory Implementation

Private Function IVectorStorageFactory_Create(ByVal length As Long) As IVectorStorage
    Set IVectorStorageFactory_Create = Me.Create(length)
End Function

Private Function IVectorStorageFactory_CreateFromRange(ByVal WorksheetRange As Range) As IVectorStorage
    Set IVectorStorageFactory_CreateFromRange = Me.CreateFromRange(WorksheetRange)
End Function
