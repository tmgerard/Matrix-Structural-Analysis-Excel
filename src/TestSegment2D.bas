Attribute VB_Name = "TestSegment2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private start_point As Point2D
Private end_point As Point2D
Private segment As Segment2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set start_point = New Point2D
    With start_point
        .x = 400
        .y = 0
    End With
    
    Set end_point = New Point2D
    With end_point
        .x = 0
        .y = 400
    End With
    
    Set segment = New Segment2D
    With segment
        Set .StartPoint = start_point
        Set .EndPoint = end_point
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
Private Sub TestLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 400 * Math.Sqr(2)

    'Act:
    Dim actual As Double
    actual = segment.Length

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Expected Error")
Private Sub TestPointAt_BadRatioValue()
    Const ExpectedError As Long = SegmentRatioError.BadValue
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim point As Point2D
    Set point = segment.PointAt(100)  ' Should be between 0 and 1

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

'@TestMethod("Basic Operation")
Private Sub TestPointAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ratio As Double
    ratio = Segment2DRatio.MakeValidRatio(0.25)
    
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 300
        .y = 100
    End With

    'Act:
    Dim actual As Point2D
    Set actual = segment.PointAt(ratio)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestPointAtMiddle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 200
        .y = 200
    End With

    'Act:
    Dim actual As Point2D
    Set actual = segment.PointAtMiddle

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestClosestPointIsStart()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 500
        .y = 20
    End With
    
    Dim expected As Point2D
    Set expected = segment.StartPoint

    'Act:
    Dim actual As Point2D
    Set actual = segment.ClosestPointTo(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestClosestPointIsEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 20
        .y = 500
    End With
    
    Dim expected As Point2D
    Set expected = segment.EndPoint

    'Act:
    Dim actual As Point2D
    Set actual = segment.ClosestPointTo(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestClosestPointIsMiddle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 250
        .y = 250
    End With
    
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 200
        .y = 200
    End With

    'Act:
    Dim actual As Point2D
    Set actual = segment.ClosestPointTo(point)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestSegmentsNoIntersection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point1 As Point2D
    Set point1 = New Point2D
    With point1
        .x = 200
        .y = 0
    End With
    
    Dim point2 As Point2D
    Set point2 = New Point2D
    With point2
        .x = 0
        .y = 200
    End With
    
    Dim otherSegment As Segment2D
    Set otherSegment = New Segment2D
    With otherSegment
        Set .StartPoint = point1
        Set .EndPoint = point2
    End With

    'Act:


    'Assert:
    Assert.IsNothing segment.IntersectionWith(otherSegment)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestSegmentsIntersection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point1 As Point2D
    Set point1 = New Point2D
    With point1
        .x = 0
        .y = 0
    End With
    
    Dim point2 As Point2D
    Set point2 = New Point2D
    With point2
        .x = 400
        .y = 400
    End With
    
    Dim otherSegment As Segment2D
    Set otherSegment = New Segment2D
    With otherSegment
        Set .StartPoint = point1
        Set .EndPoint = point2
    End With

    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 200
        .y = 200
    End With

    'Act:
    Dim actual As Point2D
    Set actual = segment.IntersectionWith(otherSegment)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

