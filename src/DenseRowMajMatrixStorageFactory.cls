VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DenseRowMajMatrixStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Matrix.Factory")

Option Explicit
Implements IMatrixStorageFactory

Private Type TDenseRowMajMatrixStorageFactory
    matrixData As IMatrixStorage
End Type

Private this As TDenseRowMajMatrixStorageFactory

'@Description("Initialize the MatrixData object")
Private Sub Class_Initialize()
    Set this.matrixData = New DenseRowMajorMatrixStorage
End Sub

'@Description("Release the MatrixData object")
Private Sub Class_Terminate()
    Set this.matrixData = Nothing
End Sub

'@Description("Create an empty matrix of a given size.")
Public Function Create(ByVal Rows As Long, ByVal Columns As Long) As DenseRowMajorMatrixStorage
    
    this.matrixData.SetSize Rows:=Rows, Columns:=Columns
    Set Create = this.matrixData
    
End Function

'@Description("Create matrix from a worksheet range")
Public Function CreateFromRange(ByVal WorksheetRange As Range) As DenseRowMajorMatrixStorage

    this.matrixData.SetSize Rows:=WorksheetRange.Rows.Count, Columns:=WorksheetRange.Columns.Count
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To this.matrixData.Rows - 1
        For column = 0 To this.matrixData.Columns - 1
            ' NOTE: Worksheet ranges are base 1 indexed
            this.matrixData.Element(row, column) = CDbl(WorksheetRange.Cells(row + 1, column + 1).Value2)
        
        Next column
    Next row
    
    Set CreateFromRange = this.matrixData

End Function

'@Description("Create identity matrix of a given size.")
Public Function CreateIdentity(ByVal Rows As Long) As DenseRowMajorMatrixStorage

    Const DIAGONAL_VALUE As Double = 1#
    
    this.matrixData.SetSize Rows:=Rows, Columns:=Rows
    
    Dim index As Long
    For index = 0 To this.matrixData.Rows - 1
        this.matrixData.Element(index, index) = DIAGONAL_VALUE
    Next index
    
    Set CreateIdentity = this.matrixData
    
End Function

Public Function CreateSquare(ByVal Rows As Long) As DenseRowMajorMatrixStorage

    this.matrixData.SetSize Rows:=Rows, Columns:=Rows
    Set CreateSquare = this.matrixData
    
End Function

' IMatrixStorageFactor Implementation

Private Function IMatrixStorageFactory_Create(ByVal Rows As Long, ByVal Columns As Long) As IMatrixStorage
    Set IMatrixStorageFactory_Create = Me.Create(Rows, Columns)
End Function

Private Function IMatrixStorageFactory_CreateFromRange(ByVal WorksheetRange As Range) As IMatrixStorage
    Set IMatrixStorageFactory_CreateFromRange = Me.CreateFromRange(WorksheetRange)
End Function

Private Function IMatrixStorageFactory_CreateIdentity(ByVal Rows As Long) As IMatrixStorage
    Set IMatrixStorageFactory_CreateIdentity = Me.CreateIdentity(Rows)
End Function

'@Description("Create an empty square matrix of a given size.")
Private Function IMatrixStorageFactory_CreateSquare(ByVal Rows As Long) As IMatrixStorage
    Set IMatrixStorageFactory_CreateSquare = Me.CreateSquare(Rows)
End Function

