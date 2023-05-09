Attribute VB_Name = "TestTrussDrawer"
'@Folder("Tests")
Option Explicit

Private Sub TestTrussDrawer()

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
        Set .Nodes = trussNodes
        Set .Bars = trussBars
    End With
    
    Dim Trans As AffineTransform
    Set Trans = New AffineTransform
    Trans.ScaleY = -1
    Trans.translateY = 100
    
    Dim drawer As TrussImager
    Set drawer = New TrussImager
    With drawer
        Set .Target = TrussDrawing
        Set .transform = Trans
    End With
    
    drawer.Draw truss

End Sub
