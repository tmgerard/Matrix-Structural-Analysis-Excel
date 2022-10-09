Attribute VB_Name = "TestVector2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private u As Vector2D
Private v As Vector2D
Private east As Vector2D
Private west As Vector2D
Private northEast As Vector2D
Private southEast As Vector2D
Private toNormalize As Vector2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set u = New Vector2D
    With u
        .u = 1
        .v = 2
    End With
    
    Set v = New Vector2D
    With v
        .u = 4
        .v = 6
    End With
    
    Set east = New Vector2D
    With east
        .u = 1
        .v = 0
    End With
    
    Set west = New Vector2D
    With west
        .u = -1
        .v = 0
    End With
    
    Set northEast = New Vector2D
    With northEast
        .u = 1
        .v = 1
    End With
    
    Set southEast = New Vector2D
    With southEast
        .u = 1
        .v = -1
    End With
    
    Set toNormalize = New Vector2D
    With toNormalize
        .u = 10
        .v = 0
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set u = Nothing
    Set v = Nothing
    Set east = Nothing
    Set west = Nothing
    Set northEast = Nothing
    Set southEast = Nothing
    Set toNormalize = Nothing
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
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 5
        .v = 8
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = u.Add(v)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestSubtract()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = -3
        .v = -4
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = u.Subtract(v)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestDotProduct()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 16
    
    'Act:
    Dim actual As Double
    actual = u.Dot(v)

    'Assert:
    Assert.IsTrue Doubles.Equal(actual, expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestCrossProduct()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = -2
    
    'Act:
    Dim actual As Double
    actual = u.Cross(v)

    'Assert:
    Assert.IsTrue Doubles.Equal(actual, expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAreParallel()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsTrue u.IsParallelTo(u)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAreNotParallel()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsFalse u.IsParallelTo(v)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestArePerpendicular()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Perpendicular As Vector2D
    Set Perpendicular = New Vector2D
    With Perpendicular
        .u = -2
        .v = 1
    End With

    'Act:

    'Assert:
    Assert.IsTrue u.IsPerpendicularTo(Perpendicular)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAreNotPerpendicular()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsFalse u.IsPerpendicularTo(v)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAngleValueZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 0

    'Act:
    Dim actual As Double
    actual = east.AngleValueTo(east)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAngleValuePi()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = WorksheetFunction.Pi

    'Act:
    Dim actual As Double
    actual = east.AngleValueTo(west)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAngleValuePiOverFourPositiveRotation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = WorksheetFunction.Pi / 4

    'Act:
    Dim actual As Double
    actual = east.AngleValueTo(northEast)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAngleValuePiOverFourNegativeRotation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = WorksheetFunction.Pi / 4

    'Act:
    Dim actual As Double
    actual = east.AngleValueTo(southEast)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestRotateZeroRadians()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = east

    'Act:
    Dim actual As Vector2D
    Set actual = east.Rotated(0)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestRotatePositiveAngle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 1 / Math.Sqr(2)
        .v = 1 / Math.Sqr(2)
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = east.Rotated(WorksheetFunction.Pi / 4)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestRotateNegativeAngle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 1 / Math.Sqr(2)
        .v = -1 / Math.Sqr(2)
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = east.Rotated(-WorksheetFunction.Pi / 4)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestRotatePiOverTwo()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 0
        .v = 1
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = east.Rotated(WorksheetFunction.Pi / 2)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestRotatePi()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = -1
        .v = 0
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = east.Rotated(WorksheetFunction.Pi)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestNormalized()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 1
        .v = 0
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = toNormalize.Normalized

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestScaledByTwo()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 2
        .v = 2
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = northEast.ScaledBy(2)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestScaledByHalf()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = 0.5
        .v = 0.5
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = northEast.ScaledBy(0.5)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestNorm()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 5
    
    Dim lengthFive As Vector2D
    Set lengthFive = New Vector2D
    With lengthFive
        .u = 3
        .v = 4
    End With

    'Act:
    Dim actual As Double
    actual = lengthFive.Norm

    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


