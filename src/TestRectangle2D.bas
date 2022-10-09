Attribute VB_Name = "TestRectangle2D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private originPt As Point2D
Private rectSize As Rectangle2DSize
Private rect As Rectangle2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set originPt = New Point2D
    With originPt
        .x = 0
        .y = 0
    End With
    
    Set rectSize = New Rectangle2DSize
    With rectSize
        .Width = 10
        .Height = 5
    End With
    
    Set rect = New Rectangle2D
    With rect
        Set .Origin = originPt
        Set .size = rectSize
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set originPt = Nothing
    Set rectSize = Nothing
    Set rect = Nothing
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
Private Sub TestContainsPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 5
        .y = 3
    End With

    'Act:

    'Assert:
    Assert.IsTrue rect.ContainsPoint(point)

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
        .x = 50
        .y = 3
    End With

    'Act:

    'Assert:
    Assert.IsFalse rect.ContainsPoint(point)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestToPolygon()
    On Error GoTo TestFail
    
    'Arrange:
    Dim points As Collection
    Set points = New Collection
    
    points.Add originPt
    
    Dim point1 As Point2D
    Set point1 = New Point2D
    With point1
        .x = 10
        .y = 0
    End With
    
    points.Add point1
    
    Dim point2 As Point2D
    Set point2 = New Point2D
    With point2
        .x = 10
        .y = 5
    End With
    
    points.Add point2
    
    Dim point3 As Point2D
    Set point3 = New Point2D
    With point3
        .x = 0
        .y = 5
    End With
    
    points.Add point3
    
    Dim expected As Polygon2D
    Set expected = New Polygon2D
    Set expected.Vertices = points

    'Act:
    Dim actual As Polygon2D
    Set actual = New Polygon2D
    
    Set actual = rect.ToPolygon

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestNoIntersectionHorizontalOverlap()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 50
        .y = 0
    End With
    
    Dim other As Rectangle2D
    Set other = New Rectangle2D
    With other
        Set .Origin = point
        Set .size = rectSize
    End With

    'Act:

    'Assert:
    Assert.IsNothing rect.IntersectsWith(other)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestNoIntersectionVerticalOverlap()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 0
        .y = 50
    End With
    
    Dim other As Rectangle2D
    Set other = New Rectangle2D
    With other
        Set .Origin = point
        Set .size = rectSize
    End With

    'Act:

    'Assert:
    Assert.IsNothing rect.IntersectsWith(other)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestIntersection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim point As Point2D
    Set point = New Point2D
    With point
        .x = 5
        .y = 2
    End With
    
    Dim other As Rectangle2D
    Set other = New Rectangle2D
    With other
        Set .Origin = point
        Set .size = rectSize
    End With
    
    Dim newSize As Rectangle2DSize
    Set newSize = New Rectangle2DSize
    With newSize
        .Width = 5
        .Height = 3
    End With
    
    Dim expected As Rectangle2D
    Set expected = New Rectangle2D
    With expected
        Set .Origin = other.Origin
        Set .size = newSize
    End With

    'Act:
    Dim actual As Rectangle2D
    Set actual = rect.IntersectsWith(other)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
