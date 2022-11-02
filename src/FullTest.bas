Attribute VB_Name = "FullTest"
'@Folder("Tests")
Option Explicit

Public Sub TestTruss()

    Dim load As Vector2D
    Set load = New Vector2D
    With load
        .u = 10
        .v = -5
    End With

    Dim trussNodes As Collection
    Set trussNodes = New Collection
    
    trussNodes.Add NodeFactory.MakeNode2D(1, Point2DFactory.MakePoint2D(0, 0), True, True)      ' Pin
    trussNodes.Add NodeFactory.MakeNode2D(2, Point2DFactory.MakePoint2D(48, 0), False, True)    ' Roller
    trussNodes.Add NodeFactory.MakeNode2D(3, Point2DFactory.MakePoint2D(48, 36), False, False)
    
    trussNodes.Item(3).AddLoad load
    
    Dim trussBars As Collection
    Set trussBars = New Collection
    
    trussBars.Add BarElement2DFactory.MakeBarElement2D(1, trussNodes.Item(1), trussNodes.Item(2), 10, 29000)
    trussBars.Add BarElement2DFactory.MakeBarElement2D(1, trussNodes.Item(2), trussNodes.Item(3), 10, 29000)
    trussBars.Add BarElement2DFactory.MakeBarElement2D(1, trussNodes.Item(1), trussNodes.Item(3), 10, 29000)
    
    Dim truss As Structure
    Set truss = New Structure
    With truss
        Set .nodes = trussNodes
        Set .Bars = trussBars
    End With
    
    Dim solution As SolutionStructure
    Set solution = truss.Solve

End Sub
