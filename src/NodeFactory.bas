Attribute VB_Name = "NodeFactory"
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Function MakeNode2D(ByRef nodeID As Long, ByRef nodeCoordinates As Point2D, _
    Optional ByRef xTrans As Boolean = True, Optional ByRef yTrans As Boolean = True, _
    Optional ByRef zRot As Boolean = True) As Node2D

    Dim Node As Node2D
    Set Node = New Node2D
    With Node
        Set .Position = nodeCoordinates
        .DOF(xTranslation) = xTrans
        .DOF(yTranslation) = yTrans
        .DOF(zRotation) = zRot
    End With
    
    Node.ID = nodeID
    
    Set MakeNode2D = Node

End Function
