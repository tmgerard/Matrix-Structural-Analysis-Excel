Attribute VB_Name = "TestLine2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private line1Base As Point2D
Private line1Direction As Vector2D
Private line1 As Line2D

Private line2Base As Point2D
Private line2Direction As Vector2D
Private line2 As Line2D

Private line3Base As Point2D
Private line3Direction As Vector2D
Private line3 As Line2D

Private line4Base As Point2D
Private line4Direction As Vector2D
Private line4 As Line2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set line1Base = New Point2D
    With line1Base
        .x = 0
        .y = 0
    End With
    
    Set line1Direction = New Vector2D
    With line1Direction
        .u = 1
        .v = 1
    End With
    
    Set line1 = New Line2D
    With line1
        Set .Base = line1Base
        Set .direction = line1Direction
    End With
    
    Set line2Base = New Point2D
    With line2Base
        .x = 10
        .y = 10
    End With
    
    Set line2Direction = New Vector2D
    With line2Direction
        .u = 1
        .v = 1
    End With
    
    Set line2 = New Line2D
    With line2
        Set .Base = line2Base
        Set .direction = line2Direction
    End With
    
    Set line3Base = New Point2D
    With line3Base
        .x = 50
        .y = 0
    End With
    
    Set line3Direction = New Vector2D
    With line3Direction
        .u = 0
        .v = 1
    End With
    
    Set line3 = New Line2D
    With line3
        Set .Base = line3Base
        Set .direction = line3Direction
    End With
    
    Set line4Base = New Point2D
    With line4Base
        .x = 0
        .y = 30
    End With
    
    Set line4Direction = New Vector2D
    With line4Direction
        .u = 1
        .v = 0
    End With
    
    Set line4 = New Line2D
    With line4
        Set .Base = line4Base
        Set .direction = line4Direction
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set line1Base = Nothing
    Set line1Direction = Nothing
    Set line1 = Nothing
    
    Set line2Base = Nothing
    Set line2Direction = Nothing
    Set line2 = Nothing
    
    Set line3Base = Nothing
    Set line3Direction = Nothing
    Set line3 = Nothing
    
    Set line4Base = Nothing
    Set line4Direction = Nothing
    Set line4 = Nothing
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
Private Sub TestParallelLinesNoIntersection()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsNothing line1.IntersectionWith(line2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestLinesIntersection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 50
        .y = 30
    End With

    'Act:
    Dim actual As Point2D
    Set actual = line3.IntersectionWith(line4)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
