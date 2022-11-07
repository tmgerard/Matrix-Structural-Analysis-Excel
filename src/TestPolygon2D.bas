Attribute VB_Name = "TestPolygon2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private point1 As Point2D
Private point2 As Point2D
Private point3 As Point2D
Private point4 As Point2D
Private polygon As Polygon2D
Private polygon2 As Polygon2D
Private polygon3 As Polygon2D
Private polygon4 As Polygon2D
Private polygon4_reverse As Polygon2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set point1 = New Point2D
    With point1
        .x = 0
        .y = 0
    End With
    
    Set point2 = New Point2D
    With point2
        .x = 30
        .y = 0
    End With
    
    Set point3 = New Point2D
    With point3
        .x = 0
        .y = 30
    End With
    
    Set point4 = New Point2D
    With point4
        .x = 30
        .y = 30
    End With
    
    Dim points As Collection
    Set points = New Collection
    points.Add point1
    points.Add point2
    points.Add point3
    
    Set polygon = New Polygon2D
    Set polygon.Vertices = points
    
    Dim points2 As Collection
    Set points2 = New Collection
    points2.Add point1
    points2.Add point2
    points2.Add point4
    
    Set polygon2 = New Polygon2D
    Set polygon2.Vertices = points2
    
    Set polygon3 = New Polygon2D
    Set polygon3.Vertices = points
    
    Dim poly4points As Collection
    Set poly4points = New Collection
    poly4points.Add MakePoint2D(0, 0)
    poly4points.Add MakePoint2D(15, 0)
    poly4points.Add MakePoint2D(15, 6)
    poly4points.Add MakePoint2D(12, 6)
    poly4points.Add MakePoint2D(12, 14)
    poly4points.Add MakePoint2D(10, 14)
    poly4points.Add MakePoint2D(10, 6)
    poly4points.Add MakePoint2D(0, 6)
    
    Set polygon4 = New Polygon2D
    Set polygon4.Vertices = poly4points
    
    Dim poly4_reversePoints As Collection
    Set poly4_reversePoints = New Collection
    poly4_reversePoints.Add MakePoint2D(0, 0)
    poly4_reversePoints.Add MakePoint2D(0, 6)
    poly4_reversePoints.Add MakePoint2D(10, 6)
    poly4_reversePoints.Add MakePoint2D(10, 14)
    poly4_reversePoints.Add MakePoint2D(12, 14)
    poly4_reversePoints.Add MakePoint2D(12, 6)
    poly4_reversePoints.Add MakePoint2D(15, 6)
    poly4_reversePoints.Add MakePoint2D(15, 0)
    
    Set polygon4_reverse = New Polygon2D
    Set polygon4_reverse.Vertices = poly4_reversePoints
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set point1 = Nothing
    Set point2 = Nothing
    Set point3 = Nothing
    Set point4 = Nothing
    Set polygon = Nothing
    Set polygon2 = Nothing
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
Private Sub TestSides()
    On Error GoTo TestFail
    
    'Arrange:
    Dim segments As Collection
    Set segments = New Collection
    
    Dim side1 As Segment2D
    Set side1 = New Segment2D
    With side1
        Set .StartPoint = point1
        Set .EndPoint = point2
    End With
    
    Dim side2 As Segment2D
    Set side2 = New Segment2D
    With side2
        Set .StartPoint = point2
        Set .EndPoint = point3
    End With
    
    Dim side3 As Segment2D
    Set side3 = New Segment2D
    With side3
        Set .StartPoint = point3
        Set .EndPoint = point1
    End With
    
    segments.Add side1
    segments.Add side2
    segments.Add side3

    'Act:
    Dim actual As Collection
    Set actual = polygon.Sides

    'Assert:
    Dim actualSegment As Segment2D
    Dim expectedSegment As Segment2D
    Dim segment As Long
    For segment = 1 To actual.count
        Set actualSegment = actual.Item(segment)
        Set expectedSegment = segments.Item(segment)
        Assert.IsTrue actualSegment.Equals(expectedSegment)
    Next segment

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAreaAddPointsCounterClockwise()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 106

    'Act:
    Dim actual As Double
    actual = polygon4.Area


    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestAreaAddPointsClockwise()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 106

    'Act:
    Dim actual As Double
    actual = polygon4_reverse.Area

    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestCentroid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 10
        .y = 10
    End With

    'Act:
    Dim actual As Point2D
    Set actual = polygon.Centroid

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestCentroidPolygon4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 8.0283
        .y = 4.0566
    End With

    'Act:
    Dim actual As Point2D
    Set actual = polygon4.Centroid

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestCentroidPolygon4_reverse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 8.0283
        .y = 4.0566
    End With

    'Act:
    Dim actual As Point2D
    Set actual = polygon4_reverse.Centroid

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestDoesNotContainPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 15
        .y = 20
    End With

    'Act:

    'Assert:
    Assert.IsFalse polygon.ContainsPoint(point)

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
        .x = 15
        .y = 10
    End With

    'Act:

    'Assert:
    Assert.IsTrue polygon.ContainsPoint(point)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestContainsVertex()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsTrue polygon.ContainsPoint(point1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestPolygonsNotEqual()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsFalse polygon.Equals(polygon2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestPolygonsEqual()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.IsTrue polygon.Equals(polygon3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
