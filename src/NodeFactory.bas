Attribute VB_Name = "NodeFactory"
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Function MakeNode2D(ByRef nodeID As Long, ByRef nodeCoordinates As Point2D, _
    Optional ByRef xConstrained As Boolean = False, Optional ByRef yConstrained As Boolean = False) As Node2D

    Dim node As Node2D
    Set node = New Node2D
    With node
        Set .Position = nodeCoordinates
        .xConstrained = xConstrained
        .yConstrained = yConstrained
    End With
    
    node.ID = nodeID
    
    Set MakeNode2D = node

End Function
