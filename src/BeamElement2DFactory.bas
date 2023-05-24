Attribute VB_Name = "BeamElement2DFactory"
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Function MakeBeamElement2D(ByRef ID As Long, ByRef StartNode As Node2D, _
    ByRef EndNode As Node2D, ByRef beamArea As Double, ByRef youngsModulus As Double, ByRef inertiaZ As Double) As BeamElement2D
    
    Dim beam As BeamElement2D
    Set beam = New BeamElement2D
    With beam
        .ID = ID
        Set .StartNode = StartNode
        Set .EndNode = EndNode
        .Material.E = youngsModulus
        .Section.Area = beamArea
        .Section.i = inertiaZ
    End With
    
    Set MakeBarElement2D = beam
    
End Function

