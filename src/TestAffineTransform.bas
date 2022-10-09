Attribute VB_Name = "TestAffineTransform"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private point As Point2D
Private scaleAT As AffineTransform
Private translateAT As AffineTransform
Private shearAT As AffineTransform

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set point = New Point2D
    With point
        .x = 2
        .y = 3
    End With
    
    Set scaleAT = New AffineTransform
    scaleAT.SetTransform 2, 5, 0, 0, 0, 0

    Set translateAT = New AffineTransform
    translateAT.SetTransform 1, 1, 10, 15, 0, 0
    
    Set shearAT = New AffineTransform
    shearAT.SetTransform 1, 1, 0, 0, 3, 4
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set point = Nothing
    Set scaleAT = Nothing
    Set translateAT = Nothing
    Set shearAT = Nothing
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
Private Sub TestScalePoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 4
        .y = 15
    End With

    'Act:
    Dim actual As Point2D
    Set actual = scaleAT.ApplyToPoint(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestTranslatePoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 12
        .y = 18
    End With

    'Act:
    Dim actual As Point2D
    Set actual = translateAT.ApplyToPoint(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestShearPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 11
        .y = 11
    End With

    'Act:
    Dim actual As Point2D
    Set actual = shearAT.ApplyToPoint(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestConcatenateScaleThenTranslate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As AffineTransform
    Set expected = New AffineTransform
    expected.SetTransform 2, 5, 10, 15, 0, 0

    'Act:
    Dim actual As AffineTransform
    Set actual = scaleAT.Concatenate(translateAT)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestConcatenateTranslateThenScale()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As AffineTransform
    Set expected = New AffineTransform
    expected.SetTransform 2, 5, 20, 75, 0, 0

    'Act:
    Dim actual As AffineTransform
    Set actual = translateAT.Concatenate(scaleAT)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestInverseTransform()
    On Error GoTo TestFail
    
    'Arrange:
    Dim trans As AffineTransform
    Set trans = New AffineTransform
    trans.SetTransform 1, 2, 3, 4, 5, 6
    
    Dim expected As AffineTransform
    Set expected = New AffineTransform

    'Act:
    Dim actual As AffineTransform
    Set actual = trans.Concatenate(trans.Inverse)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
