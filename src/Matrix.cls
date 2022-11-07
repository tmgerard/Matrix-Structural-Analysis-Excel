VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Equations")
Option Explicit

Private Type TMatrix
    numRows As Long
    numColumns As Long
    Data() As Double    ' dense row-major storage scheme
End Type

Private this As TMatrix

Public Enum MatrixErrors
    BadRows = 500 + vbObjectError
    BadColumns
    MatrixSizeMismatch
End Enum

'@Description "Sets the number of rows and columns in the matrix."
Public Function SetSize(ByRef numRows As Long, ByRef numColumns As Long) As Matrix
Attribute SetSize.VB_Description = "Sets the number of rows and columns in the matrix."

    If numRows <= 0 Then
        Err.Raise Number:=MatrixErrors.BadRows, _
                  source:="Matrix.SetSize", _
                  Description:="Row size cannot be zero or negative."
    End If
    
    If numColumns <= 0 Then
        Err.Raise Number:=MatrixErrors.BadColumns, _
                  source:="Matrix.SetSize", _
                  Description:="Column size cannot be zero or negative."
    End If

    With this
        .numRows = numRows
        .numColumns = numColumns
        ReDim .Data(numRows * numColumns - 1)   ' storing matrix in a one-dimensional array
    End With
    
    Set SetSize = Me
    
End Function

'@Description "Returns the number of rows in the matrix."
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns the number of rows in the matrix."
    Rows = this.numRows
End Property

'@Description "Returns the number of columns in the matrix."
Public Property Get Columns() As Long
Attribute Columns.VB_Description = "Returns the number of columns in the matrix."
    Columns = this.numColumns
End Property

'@Description "Returns the value at the given row and column in the matrix."
Public Property Get ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long) As Double
Attribute ValueAt.VB_Description = "Returns the value at the given row and column in the matrix."
    ValueAt = this.Data(rowIndex * this.numColumns + columnIndex)
End Property

'@Description "Assigns a value at the given row and column in the matrix."
Public Property Let ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long, ByRef value As Double)
Attribute ValueAt.VB_Description = "Assigns a value at the given row and column in the matrix."
    this.Data(rowIndex * this.numColumns + columnIndex) = value
End Property

'@Description "Returns the value at the given row and column of the transposed matrix."
Public Property Get ValueAtTransposed(ByRef rowIndex As Long, ByRef columnIndex As Long) As Double
Attribute ValueAtTransposed.VB_Description = "Returns the value at the given row and column of the transposed matrix."
    ValueAtTransposed = Me.ValueAt(columnIndex, rowIndex)
End Property

'@Description "Returns True if the matrix row and column counts are equal."
Public Function IsSquare() As Boolean
Attribute IsSquare.VB_Description = "Returns True if the matrix row and column counts are equal."
    IsSquare = (this.numColumns = this.numRows)
End Function

'@Description "Add a given value to a specified value in the Matrix."
Public Function AddTo(ByRef rowIndex As Long, ByRef columnIndex As Long, ByRef value As Double) As Matrix
Attribute AddTo.VB_Description = "Add a given value to a specified value in the Matrix."
    Me.ValueAt(rowIndex, columnIndex) = Me.ValueAt(rowIndex, columnIndex) + value
    Set AddTo = Me
End Function

'@Description "Populates Matrix object with the values of a given arrray."
Public Function SetData(ByRef matrixData() As Double) As Matrix
Attribute SetData.VB_Description = "Populates Matrix object with the values of a given arrray."
    ' matrixData array must be stored in row-major format
    If Not ArrayLength(matrixData) = (this.numColumns * this.numRows) Then
        Err.Raise Number:=MatrixErrors.MatrixSizeMismatch, _
                  source:="Matrix.SetData", _
                  Description:="Cannot set data: Size mismatch"
    End If
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = matrixData(index)
    Next index
    
    Set SetData = Me
End Function

Public Function SetIdentityRow(ByRef rowIndex As Long) As Matrix

    Dim columnIndex As Long
    For columnIndex = 0 To this.numColumns - 1
        If rowIndex = columnIndex Then
            Me.ValueAt(rowIndex, columnIndex) = 1
        Else
            Me.ValueAt(rowIndex, columnIndex) = 0
        End If
    Next columnIndex
    
    Set SetIdentityRow = Me
End Function

Public Function SetIdentityColumn(ByRef columnIndex As Long) As Matrix

    Dim rowIndex As Long
    For rowIndex = 0 To this.numRows - 1
        If rowIndex = columnIndex Then
            Me.ValueAt(rowIndex, columnIndex) = 1
        Else
            Me.ValueAt(rowIndex, columnIndex) = 0
        End If
    Next rowIndex
    
    Set SetIdentityColumn = Me
End Function

Public Function ScaleValues(ByRef factor As Double) As Matrix
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = factor * this.Data(index)
    Next index
    
    Set ScaleValues = Me
End Function


Public Function Equals(ByRef other As Matrix) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Matrix Then ' other is not a Matrix
        Equals = False
    ElseIf Not this.numRows = other.Rows Then
        Equals = False
    ElseIf Not this.numColumns = other.Columns Then
        Equals = False
    Else
        Dim rowIndex As Long
        Dim columnIndex As Long
        For rowIndex = 0 To this.numRows - 1
            For columnIndex = 0 To this.numColumns - 1
                If Not Doubles.Equal(Me.ValueAt(rowIndex, columnIndex), other.ValueAt(rowIndex, columnIndex)) Then
                    Equals = False
                    Exit Function
                End If
            Next columnIndex
        Next rowIndex
        Equals = True
    End If

End Function
