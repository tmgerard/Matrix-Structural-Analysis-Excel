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
Implements IVectorStorageFactory

' TODO: Implement IMatrixStorageFactory

Private Type TDenseVectoryStorageFactory
    VectorData As DenseColumnVectorStorage
End Type

Private this As TDenseVectoryStorageFactory
Private Const CLASS_NAME As String = "DenseVectorStorageFactory"

Private Sub Class_Initialize()
    Set this.VectorData = New DenseColumnVectorStorage
End Sub

Private Sub Class_Terminate()
    Set this.VectorData = Nothing
End Sub

'@Desctiption("Create an empty vector of a given length.")
Public Function Create(ByVal Length As Long) As DenseColumnVectorStorage
    this.VectorData.Length = Length
    Set Create = this.VectorData
End Function

'@Description("Create vector from a worksheet range")
Public Function CreateFromRange(ByVal WorksheetRange As Range) As DenseColumnVectorStorage
    
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
Private Function VectorFromRowRange(ByVal WorksheetRange As Range) As DenseColumnVectorStorage
    
    Dim rangeRow As Long
    rangeRow = WorksheetRange.row
    
    Dim Columns As Long
    Columns = WorksheetRange.Columns.Count
    
    this.VectorData.Length = Columns
    
    Dim column As Long
    For column = 0 To Columns - 1
        this.VectorData.Element(column) = CDbl(WorksheetRange.Cells(rangeRow, column + 1).Value2)
    Next column
    
    Set VectorFromRowRange = this.VectorData
    
End Function

'@Description("Create vector from a column oriented worksheet range")
Private Function VectorFromColumnRange(ByVal WorksheetRange As Range) As DenseColumnVectorStorage

    Dim Rows As Long
    Rows = WorksheetRange.Rows.Count
    
    Dim rangeColumn As Long
    rangeColumn = WorksheetRange.column
    
    this.VectorData.Length = Rows
    
    Dim row As Long
    For row = 0 To Rows - 1
        this.VectorData.Element(row) = CDbl(WorksheetRange.Cells(row + 1, rangeColumn).Value2)
    Next row
    
    Set VectorFromColumnRange = this.VectorData

End Function

' IVectorStorageFactory Implementation

Private Function IVectorStorageFactory_Create(ByVal Length As Long) As IVectorStorage
    Set IVectorStorageFactory_Create = Me.Create(Length)
End Function

Private Function IVectorStorageFactory_CreateFromRange(ByVal WorksheetRange As Range) As IVectorStorage
    Set IVectorStorageFactory_CreateFromRange = Me.CreateFromRange(WorksheetRange)
End Function
