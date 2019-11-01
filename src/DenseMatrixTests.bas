Attribute VB_Name = "DenseMatrixTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests.Matrix")

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

'@TestMethod("Matrix Operation")
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 2
    Const COLUMNS As Long = 2
    Const EXPECTEDVALUE As Double = 3
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, COLUMNS)
        .Element(0, 0) = 1
        .Element(0, 1) = 1
        .Element(1, 0) = 1
        .Element(1, 1) = 1
    End With
    
    Dim MatrixB As DenseMatrix
    Set MatrixB = New DenseMatrix
    
    With MatrixB
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, COLUMNS)
        .Element(0, 0) = 2
        .Element(0, 1) = 2
        .Element(1, 0) = 2
        .Element(1, 1) = 2
    End With

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = MatrixA.Add(MatrixB)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, MatrixC.Element(0, 0)
    Assert.AreEqual EXPECTEDVALUE, MatrixC.Element(0, 1)
    Assert.AreEqual EXPECTEDVALUE, MatrixC.Element(1, 0)
    Assert.AreEqual EXPECTEDVALUE, MatrixC.Element(1, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestAddMismatchedMatrixDimensions()
    Const ExpectedError As Long = MatrixError.Addition
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 2
    Const COLUMNS As Long = 2
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, COLUMNS)
    End With
    
    Dim MatrixB As DenseMatrix
    Set MatrixB = New DenseMatrix
    
    With MatrixB
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows + 1, COLUMNS + 1)
    End With

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = MatrixA.Add(MatrixB)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestIsSquareTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 2
    Const COLUMNS As Long = 2
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, COLUMNS)
    End With

    'Act:

    'Assert:
    Assert.IsTrue MatrixA.IsSquare

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestIsSquareFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 2
    Const COLUMNS As Long = 4
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, COLUMNS)
    End With

    'Act:

    'Assert:
    Assert.IsFalse MatrixA.IsSquare

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestMultiply()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(3, 2)
        .Element(0, 0) = 5
        .Element(0, 1) = 3
        .Element(1, 0) = 4
        .Element(1, 1) = 1
        .Element(2, 0) = -2
        .Element(2, 1) = 8
    End With
    
    Dim MatrixB As DenseMatrix
    Set MatrixB = New DenseMatrix
    
    With MatrixB
        .MatrixStorage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(2, 2)
        .Element(0, 0) = 2
        .Element(0, 1) = 7
        .Element(1, 0) = -3
        .Element(1, 1) = 4
    End With

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = MatrixA.Multiply(MatrixB)
    
    Dim output As IMatrixOutput
    Set output = New DenseMatrixOutput
    
    Debug.Print
    output.ToImmediateWindow MatrixA.MatrixStorage
    Debug.Print
    output.ToImmediateWindow MatrixB.MatrixStorage
    Debug.Print
    output.ToImmediateWindow MatrixC.MatrixStorage

    'Assert:
    Assert.AreEqual 1#, MatrixC.Element(0, 0)
    Assert.AreEqual 47#, MatrixC.Element(0, 1)
    Assert.AreEqual 5#, MatrixC.Element(1, 0)
    Assert.AreEqual 32#, MatrixC.Element(1, 1)
    Assert.AreEqual -28#, MatrixC.Element(2, 0)
    Assert.AreEqual 18#, MatrixC.Element(2, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestMultiplyMismatchedSizes()
    Const ExpectedError As Long = MatrixError.Multiplication
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property")
Private Sub TestGetRows()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestScalarMultiply()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestSubtraction()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestSubtractionMismatchedSizes()
    Const ExpectedError As Long = MatrixError.Subtraction
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestSwapRows()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestTranspose()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

