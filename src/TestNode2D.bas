Attribute VB_Name = "TestNode2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Model.Structure Model")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private point As Point2D
Private load1 As Vector2D
Private load2 As Vector2D
Dim Node As Node2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set point = New Point2D
    With point
        .x = 2
        .y = 5
    End With
    
    Set load1 = New Vector2D
    With load1
        .u = 10
        .v = 20
    End With
    
    Set load2 = New Vector2D
    With load2
        .u = 30
        .v = 40
    End With
    
    Set Node = MakeNode2D(1, point)
    Node.AddLoad load1
    Node.AddLoad load2
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set point = Nothing
    Set load1 = Nothing
    Set load2 = Nothing
    Set Node = Nothing
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
Private Sub TestNetLoad()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 40
        .v = 60
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = Node.NetLoad
    

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
