Attribute VB_Name = "VectorOperatorTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.LinearAlgebra.Vector")
'@IgnoreModule

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private operator As VectorOperator

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
    
    Set operator = New VectorOperator
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Vector Operation")
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 2
    
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseRowVectorStub

    'Act:
    Dim vectorC As DenseVector
    Set vectorC = operator.Add(vectorA, vectorB)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(0)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(1)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestAddMisMatchedOrientation()
    Const ExpectedError As Long = VectorError.Addition
    On Error GoTo TestFail
    
    'Arrange:
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseColumnVectorStub

    'Act:
    Dim vectorC As DenseVector
    Set vectorC = operator.Add(vectorA, vectorB)

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

'@TestMethod("Vector Operation")
Private Sub TestSubtract()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 0
    
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseRowVectorStub

    'Act:
    Dim vectorC As DenseVector
    Set vectorC = operator.Subtract(vectorA, vectorB)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(0)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(1)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestSubtractMisMatchedOrientation()
    Const ExpectedError As Long = VectorError.Subtraction
    On Error GoTo TestFail
    
    'Arrange:
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseColumnVectorStub

    'Act:
    Dim vectorC As DenseVector
    Set vectorC = operator.Subtract(vectorA, vectorB)

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

'@TestMethod("Vector Operation")
Private Sub TestDotProduct()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 3
    
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseRowVectorStub

    'Act:
    Dim actual As Double
    actual = operator.DotProduct(vectorA, vectorB)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Vector Operation")
Private Sub TestScalarMultiply()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 3
    
    Dim vectorA As IVector
    Set vectorA = New DenseRowVectorStub

    'Act:
    Dim vectorC As IVector
    Set vectorC = operator.ScalarMultiply(vectorA, EXPECTEDVALUE)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(0)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(1)
    Assert.AreEqual EXPECTEDVALUE, vectorC.Element(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Vector Operation")
Private Sub TestCrossProduct()
    On Error GoTo TestFail
    
    'Arrange:
    Dim vectorA As IVector
    Set vectorA = New DenseVectorAStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseVectorBStub

    'Act:
    Dim vectorC As IVector
    Set vectorC = operator.CrossProduct(vectorA, vectorB)

    'Assert:
    Assert.AreEqual -15#, vectorC.Element(0)
    Assert.AreEqual -2#, vectorC.Element(1)
    Assert.AreEqual 39#, vectorC.Element(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestCrossProductVectorLengthNotThree()
    Const ExpectedError As Long = VectorError.CrossProduct
    On Error GoTo TestFail
    
    'Arrange:
    Dim vectorA As IVector
    Set vectorA = New DenseVectorXStub
    
    Dim vectorB As IVector
    Set vectorB = New DenseVectorXStub

    'Act:
    Dim vectorC As IVector
    Set vectorC = operator.CrossProduct(vectorA, vectorB)

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

'@TestMethod("Vector Operation")
Private Sub TestEuclideanDistance()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTEDVALUE As Double = 2#
    Dim vectorA As IVector
    Set vectorA = New DenseVectorXStub

    'Act:
    Dim Distance As Double
    Distance = operator.EuclideanDistance(vectorA)

    'Assert:
    Assert.AreEqual EXPECTEDVALUE, Distance

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
