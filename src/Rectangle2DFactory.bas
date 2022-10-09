Attribute VB_Name = "Rectangle2DFactory"
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Public Enum RectangleFactoryErrors
    NoPoints = 400 + vbObjectError
    BadPointObject
End Enum

Public Function MakeRectangleContaining(ByRef points As Collection) As Rectangle2D

    If points.count = 0 Then
        Err.Raise Number:=RectangleFactoryErrors.NoPoints, _
                  Source:="Rectangle2DFactory.MakeRectangleContaining", _
                  Description:="Expected at least one point"
    End If
    
    Dim point As Point2D
    For Each point In points
        If Not TypeOf point Is Point2D Then
            Err.Raise Number:=RectangleFactoryErrors.BadPointObject, _
                      Source:="Rectangle2DFactory.MakeRectangleContaining", _
                      Description:="All objects must be of type Point2D"
        End If
    Next point
    
    Dim firstPoint As Point2D
    Set firstPoint = points.Item(1)
    
    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    minX = firstPoint.x
    maxX = firstPoint.x
    minY = firstPoint.y
    maxY = firstPoint.y
    
    Dim index As Long
    For index = 2 To points.count
        minX = WorksheetFunction.MIN(minX, points.Item(index).x)
        maxX = WorksheetFunction.MAX(maxX, points.Item(index).x)
        minY = WorksheetFunction.MIN(minY, points.Item(index).y)
        maxY = WorksheetFunction.MAX(maxY, points.Item(index).y)
    Next index
    
    Dim originPoint As Point2D
    Set originPoint = New Point2D
    With originPoint
        .x = minX
        .y = minY
    End With
    
    Dim rectSize As Rectangle2DSize
    Set rectSize = New Rectangle2DSize
    With rectSize
        .Width = maxX - minX
        .Height = maxY - minY
    End With
    
    Dim newRect As Rectangle2D
    Set newRect = New Rectangle2D
    With newRect
        Set .Origin = originPoint
        Set .size = rectSize
    End With
    
    Set MakeRectangleContaining = newRect

End Function

Public Function MakeRectangleContainingWithMargins(ByRef points As Collection, ByRef margin As Double)

    Dim rect As Rectangle2D
    Set rect = MakeRectangleContaining(points)
    
    Dim newOrigin As Point2D
    Set newOrigin = New Point2D
    With newOrigin
        .x = rect.Origin.x - margin
        .y = rect.Origin.y - margin
    End With
    
    Dim newSize As Rectangle2DSize
    Set newSize = New Rectangle2DSize
    With newSize
        .Width = 2 * margin + rect.size.Width
        .Height = 2 * margin + rect.size.Height
    End With
    
    Dim biggerRect As Rectangle2D
    Set biggerRect = New Rectangle2D
    With biggerRect
        Set .Origin = newOrigin
        Set .size = newSize
    End With
    
    Set MakeRectangleContainingWithMargins = biggerRect

End Function
