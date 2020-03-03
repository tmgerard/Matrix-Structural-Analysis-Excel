Attribute VB_Name = "DenseMatrixTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests.LinearAlgebra.Matrix")
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

'@TestMethod("Matrix Operation")
Private Sub TestIsSquareTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    Dim storageStub As IMatrixStorage
    Set storageStub = New SquareDenseStorageStub
    
    With MatrixA
        Set .Storage = storageStub
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
    Const Columns As Long = 4
    
    Dim MatrixA As DenseMatrix
    Set MatrixA = New DenseMatrix
    
    Dim storageStub As IMatrixStorage
    Set storageStub = New NotSquareDenseStorageStub
    
    With MatrixA
        Set .Storage = storageStub
    End With

    'Act:

    'Assert:
    Assert.IsFalse MatrixA.IsSquare

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


