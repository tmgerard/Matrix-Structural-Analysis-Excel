Attribute VB_Name = "TestCholesky"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Equations")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private solver As Cholesky

Private sysMatrix As Matrix
Private sysData(15) As Double

Private lowMatrix As Matrix
Private lowData(15) As Double

Private sysVector As Vector
Private sysVecData(3) As Double

Private lowSolution As Vector
Private lowSolutionData(3) As Double

Private solution As Vector
Private solutionData(3) As Double


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    sysData(0) = 4
    sysData(1) = -2
    sysData(2) = 4
    sysData(3) = 2
    sysData(4) = -2
    sysData(5) = 10
    sysData(6) = -2
    sysData(7) = -7
    sysData(8) = 4
    sysData(9) = -2
    sysData(10) = 8
    sysData(11) = 4
    sysData(12) = 2
    sysData(13) = -7
    sysData(14) = 4
    sysData(15) = 7
    
    Set sysMatrix = New Matrix
    Set sysMatrix = sysMatrix.SetSize(4, 4).SetData(sysData)
    
    lowData(0) = 2
    lowData(1) = 0
    lowData(2) = 0
    lowData(3) = 0
    lowData(4) = -1
    lowData(5) = 3
    lowData(6) = 0
    lowData(7) = 0
    lowData(8) = 2
    lowData(9) = 0
    lowData(10) = 2
    lowData(11) = 0
    lowData(12) = 1
    lowData(13) = -2
    lowData(14) = 1
    lowData(15) = 1
    
    Set lowMatrix = New Matrix
    Set lowMatrix = lowMatrix.SetSize(4, 4).SetData(lowData)
    
    Set solver = New Cholesky

    sysVecData(0) = 20
    sysVecData(1) = -16
    sysVecData(2) = 40
    sysVecData(3) = 28
    
    Set sysVector = New Vector
    Set sysVector = sysVector.SetLength(4).SetData(sysVecData)
    
    lowSolutionData(0) = 10
    lowSolutionData(1) = -2
    lowSolutionData(2) = 10
    lowSolutionData(3) = 4
    
    Set lowSolution = New Vector
    Set lowSolution = lowSolution.SetLength(4).SetData(lowSolutionData)
    
    solutionData(0) = 1
    solutionData(1) = 2
    solutionData(2) = 3
    solutionData(3) = 4
    
    Set solution = New Vector
    Set solution = solution.SetLength(4).SetData(solutionData)
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set sysMatrix = Nothing
    Erase sysData
    
    Set lowMatrix = Nothing
    Erase lowData
    
    Set solver = Nothing
    
    Set sysVector = Nothing
    Erase sysVecData
    
    Set lowSolution = Nothing
    Erase lowSolutionData
    
    Set solution = Nothing
    Erase solutionData
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Algorithm")
Private Sub TestLowerMatrixDecomposition()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Matrix
    Set actual = solver.LowDecomposition(sysMatrix)

    'Assert:
    Assert.IsTrue lowMatrix.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestLowerForwardSubstitution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.ForwardSubstitution(lowMatrix, sysVector)

    'Assert:
    Assert.IsTrue lowSolution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestUpperBackSubstitution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.BackSubstutution(lowMatrix, lowSolution)

    'Assert:
    Assert.IsTrue solution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestSolution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.Solve(sysMatrix, sysVector)

    'Assert:
    Assert.IsTrue solution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
