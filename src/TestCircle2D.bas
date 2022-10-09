Attribute VB_Name = "TestCircle2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private originPt As Point2D
Private circ As Circle2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set originPt = New Point2D
    With originPt
        .x = 10
        .y = 10
    End With
    
    Set circ = New Circle2D
    With circ
        Set .center = originPt
        .radius = 10
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Basic Operation")
Private Sub TestArea()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = WorksheetFunction.Pi * 100

    'Act:

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, circ.Area)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestCircumference()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = WorksheetFunction.Pi * 20

    'Act:

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, circ.Circumference)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestContainsPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 11
        .y = 12
    End With

    'Act:

    'Assert:
    Assert.IsTrue circ.ContainsPoint(point)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestDoesNotContainsPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 110
        .y = 12
    End With

    'Act:

    'Assert:
    Assert.IsFalse circ.ContainsPoint(point)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
