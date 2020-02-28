Attribute VB_Name = "LupSolverTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.LinearAlgebra.Solver")
'@IgnoreModule

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private solver As LupSolver

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
    
    Set solver = New LupSolver
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
    
    Set solver = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Linear Solver")
Private Sub TestSolve()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As IVector
    Set Expected = New UnkownVectorXStub
    
    Dim a As IMatrix
    Set a = New CoefficientMatrixAStub
    
    Dim b As IVector
    Set b = New SolutionVectorBStub
    
    Dim X As IVector

    'Act:
    Set X = solver.Solve(a, b)

    'Assert:
    Dim i As Long
    For i = 0 To X.length - 1
        Assert.AreEqual Expected.Element(i), X.Element(i)
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestSolveSingularMatrix()
    Const ExpectedError As Long = SolverError.SingularMatrix
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As IVector
    Set Expected = New UnkownVectorXStub
    
    Dim a As IMatrix
    Set a = New SingluarCeofficientMatrixAStub
    
    Dim b As IVector
    Set b = New SolutionVectorBStub
    
    Dim X As IVector

    'Act:
    Set X = solver.Solve(a, b)

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

