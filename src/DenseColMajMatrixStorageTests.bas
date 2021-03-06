Attribute VB_Name = "DenseColMajMatrixStorageTests"
Option Explicit
Option Private Module

Private Const CREATE_ROWS As Long = 5
Private Const CREATE_COLUMNS As Long = 6
Private Const BAD_ROWS As Long = 0
Private Const BAD_COLUMNS As Long = 0
Private Const ELEMENT_VALUE As Double = 5#

'@TestModule
'@Folder("Tests.LinearAlgebra.MatrixStorage")
'@IgnoreModule

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        'Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New AssertClass
        'Set Fakes = New FakesProvider
    #End If
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Operation")
Public Sub TestSetSize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    matrix.SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS

    'Assert:
    Assert.AreEqual CREATE_ROWS, matrix.Rows
    Assert.AreEqual CREATE_COLUMNS, matrix.Columns

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Public Sub TestSetSizeBadRowSize()
    Const ExpectedError As Long = MatrixError.RowSize
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=BAD_ROWS, Columns:=CREATE_COLUMNS
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestSetSizeBadColumnSize()
    Const ExpectedError As Long = MatrixError.ColumnSize
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=BAD_COLUMNS
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property")
Public Sub TestElementProperty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS
        .Element(0, 0) = ELEMENT_VALUE
    End With

    'Assert:
    Assert.AreEqual ELEMENT_VALUE, matrix.Element(0, 0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Public Sub TestElementPropertyBadRowIndex()
    Const ExpectedError As Long = MatrixError.RowRange
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS
        .Element(CREATE_ROWS, 0) = ELEMENT_VALUE
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestElementPropertyBadColumnIndex()
    Const ExpectedError As Long = MatrixError.ColumnRange
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS
        .Element(0, CREATE_COLUMNS) = ELEMENT_VALUE
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestElementPropertyBadColumnIndexNegative()
    Const ExpectedError As Long = MatrixError.ColumnRange
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseRowMajorMatrixStorage
    Set matrix = New DenseRowMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS
        .Element(0, -1) = ELEMENT_VALUE
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestElementPropertyBadRowIndexNegative()
    Const ExpectedError As Long = MatrixError.RowRange
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS
        .Element(-1, CREATE_COLUMNS) = ELEMENT_VALUE
    End With

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Assert.Succeed
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property")
Public Sub TestGetRows()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    matrix.SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS

    'Assert:
    Assert.AreEqual CREATE_ROWS, matrix.Rows

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestGetColumns()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    matrix.SetSize Rows:=CREATE_ROWS, Columns:=CREATE_COLUMNS

    'Assert:
    Assert.AreEqual CREATE_COLUMNS, matrix.Columns

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Operation")
Public Sub TestClear()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As IMatrixStorage
    Set matrix = New SquareDenseStorageStub

    'Act:
    matrix.Clear

    'Assert:
    Dim row As Long
    Dim column As Long
    For row = 0 To matrix.Rows - 1
        For column = 0 To matrix.Columns - 1
            Assert.AreEqual 0#, matrix.Element(row, column)
        Next column
    Next row

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Operation")
Public Sub TestClone()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrix As DenseColumnMajorMatrixStorage
    Set matrix = New DenseColumnMajorMatrixStorage

    'Act:
    With matrix
        .SetSize Rows:=2, Columns:=2
        .Element(0, 0) = 1#
        .Element(0, 1) = 1#
        .Element(1, 0) = 1#
        .Element(1, 1) = 1#
    End With
    
    Dim newMatrix As DenseColumnMajorMatrixStorage
    Set newMatrix = matrix.Clone
    
    'Assert:
    Assert.IsTrue Not ObjPtr(matrix) = ObjPtr(newMatrix)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

