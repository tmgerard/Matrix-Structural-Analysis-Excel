Attribute VB_Name = "MatrixOperatorTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Linear Algebra.Matrix")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private operator As MatrixOperator

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        'Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New AssertClass
        'Set Fakes = New FakesProvider
    #End If
    
    Set operator = New MatrixOperator
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
    Set operator = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 3
    
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixAStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixBStub
    
    'Act:
    Dim MatrixC As IMatrix
    Set MatrixC = operator.Add(MatrixA, MatrixB)

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
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixAStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixXStub

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Add(MatrixA, MatrixB)

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
Private Sub TestMultiply()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixXStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixYStub

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Multiply(MatrixA, MatrixB)

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
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixXStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixYStub

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Multiply(MatrixB, MatrixA)

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
Private Sub TestScalarMultiply()
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 2
    Const Columns As Long = 2
    Const EXPECTED_VALUE As Double = 3
    
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixAStub

    'Act:
    Dim MatrixC As IMatrix
    Set MatrixC = operator.ScalarMultiply(MatrixA, EXPECTED_VALUE)

    'Assert:
    Assert.AreEqual EXPECTED_VALUE, MatrixC.Element(0, 0)
    Assert.AreEqual EXPECTED_VALUE, MatrixC.Element(0, 1)
    Assert.AreEqual EXPECTED_VALUE, MatrixC.Element(1, 0)
    Assert.AreEqual EXPECTED_VALUE, MatrixC.Element(1, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestSubtraction()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = -1
    
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixAStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixBStub

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Subtract(MatrixA, MatrixB)

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
Private Sub TestSubtractionMismatchedSizes()
    Const ExpectedError As Long = MatrixError.Subtraction
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixAStub
    
    Dim MatrixB As IMatrix
    Set MatrixB = New DenseMatrixXStub

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Subtract(MatrixA, MatrixB)

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
    Const ROW_ONE_EXPECTED_VALUE As Double = 2#
    Const ROW_TWO_EXPECTED_VALUE As Double = 1#
    
    Dim MatrixA As IMatrix
    Set MatrixA = New DenseMatrixZStub

    'Act:
    operator.SwapRows matrix:=MatrixA, firstRowIndex:=0, secondRowIndex:=1

    'Assert:
    Assert.AreEqual ROW_ONE_EXPECTED_VALUE, MatrixA.Element(0, 0)
    Assert.AreEqual ROW_ONE_EXPECTED_VALUE, MatrixA.Element(0, 1)
    Assert.AreEqual ROW_TWO_EXPECTED_VALUE, MatrixA.Element(1, 0)
    Assert.AreEqual ROW_TWO_EXPECTED_VALUE, MatrixA.Element(1, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Matrix Operation")
Private Sub TestTranspose()
    On Error GoTo TestFail
    
    'Arrange:
    Const Rows As Long = 3
    Const Columns As Long = 2
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    With MatrixA
        .Storage = MatrixStorageFactory.CreateFactory(DenseColumnMajor).Create(Rows, Columns)
        .Element(0, 0) = 1#
        .Element(0, 1) = 2#
        .Element(1, 0) = 3#
        .Element(1, 1) = 4#
        .Element(2, 0) = 5#
        .Element(2, 1) = 6#
    End With

    'Act:
    Dim MatrixC As DenseMatrix
    Set MatrixC = operator.Transpose(MatrixA)

    'Assert:
    Assert.AreEqual 1#, MatrixC.Element(0, 0)
    Assert.AreEqual 3#, MatrixC.Element(0, 1)
    Assert.AreEqual 5#, MatrixC.Element(0, 2)
    Assert.AreEqual 2#, MatrixC.Element(1, 0)
    Assert.AreEqual 4#, MatrixC.Element(1, 1)
    Assert.AreEqual 6#, MatrixC.Element(1, 2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

