Attribute VB_Name = "TestBarElement2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Model.Structure Model")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private barArea As Double
Private modulus As Double
Private nodeA As Node2D
Private nodeB As Node2D
Private bar As BarElement2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set nodeA = MakeNode2D(1, MakePoint2D(0, 0))
    Set nodeB = MakeNode2D(2, MakePoint2D(2, 1))
    barArea = Math.Sqr(5)
    modulus = 5
    Set bar = MakeBarElement2D(1, nodeA, nodeB, barArea, modulus)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set nodeA = Nothing
    Set nodeB = Nothing
    Set bar = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Stiffness Matrix")
Private Sub TestGlobalStiffnessMatrix()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matData(15) As Double
    matData(0) = 4
    matData(1) = 2
    matData(2) = -4
    matData(3) = -2
    matData(4) = 2
    matData(5) = 1
    matData(6) = -2
    matData(7) = -1
    matData(8) = -4
    matData(9) = -2
    matData(10) = 4
    matData(11) = 2
    matData(12) = -2
    matData(13) = -1
    matData(14) = 2
    matData(15) = 1
    
    Dim expected As Matrix
    Set expected = New Matrix
    Set expected = expected.SetSize(4, 4).SetData(matData)

    'Act:
    Dim actual As Matrix
    Set actual = bar.GlobalStiffnessMatrix

    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
