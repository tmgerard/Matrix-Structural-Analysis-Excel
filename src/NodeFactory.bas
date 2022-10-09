Attribute VB_Name = "NodeFactory"
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Function MakeNode2D(ByRef nodeID As Long, ByRef nodeCoordinates As Point2D, _
    Optional ByRef xConstrained As Boolean = False, Optional ByRef yConstrained As Boolean = False) As Node2D

    Dim Node As Node2D
    Set Node = New Node2D
    With Node
        Set .Position = nodeCoordinates
        .xConstrained = xConstrained
        .yConstrained = yConstrained
    End With
    
    Node.ID = nodeID
    
    Set MakeNode2D = Node

End Function
