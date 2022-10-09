Attribute VB_Name = "BarElement2DFactory"
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Function MakeBarElement2D(ByRef ID As Long, ByRef StartNode As Node2D, _
    ByRef EndNode As Node2D, ByRef barArea As Double, ByRef youngsModulus As Double) As BarElement2D
    
    Dim bar As BarElement2D
    Set bar = New BarElement2D
    With bar
        .ID = ID
        Set .StartNode = StartNode
        Set .EndNode = EndNode
        .Area = barArea
        .ElasticModulus = youngsModulus
    End With
    
    Set MakeBarElement2D = bar
    
End Function

