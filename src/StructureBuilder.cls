VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructureBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Input")
Option Explicit

Private Type TStructureBuilder
    CrossSections As Collection
    nodes As Collection
    Materials As Collection
    Members As Collection
    CrossSectionTable As ListObject
    NodeTable As ListObject
    MaterialTable As ListObject
    MemberTable As ListObject
    LoadTable As ListObject
End Type

Private this As TStructureBuilder

Private Sub Class_Initialize()
    With this
        Set .CrossSections = New Collection
        Set .nodes = New Collection
        Set .Materials = New Collection
        Set .Members = New Collection
        Set .CrossSectionTable = TrussInput.ListObjects(InputSettings.CrossSectionTableName)
        Set .NodeTable = TrussInput.ListObjects(InputSettings.NodesTableName)
        Set .MaterialTable = TrussInput.ListObjects(InputSettings.MaterialsTableName)
        Set .MemberTable = TrussInput.ListObjects(InputSettings.MembersTableName)
        Set .LoadTable = TrussInput.ListObjects(InputSettings.LoadsTableName)
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .CrossSections = Nothing
        Set .nodes = Nothing
        Set .Materials = Nothing
        Set .Members = Nothing
        Set .CrossSectionTable = Nothing
        Set .NodeTable = Nothing
        Set .MaterialTable = Nothing
        Set .MemberTable = Nothing
        Set .LoadTable = Nothing
    End With
End Sub

Public Function Build() As Structure

    If this.nodes Is Nothing Or this.Members Is Nothing Then
        ' TODO: raise error if no loads or members defined
    End If
    
    CheckTableForNull this.NodeTable
    CheckTableForNull this.MemberTable
    CheckTableForNull this.MaterialTable
    CheckTableForNull this.LoadTable
    
    GetNodes
    GetCrossSections
    GetMaterials
    GetCrossSections
    GetMembers
    GetLoads
    
    Dim struct As Structure
    Set struct = New Structure
    With struct
        Set .nodes = this.nodes
        Set .Bars = this.Members
    End With
    
    Set Build = struct
    
End Function

Private Sub GetNodes()
    Dim xConstrained As Boolean
    Dim yConstrained As Boolean
    
    Dim row As Long
    For row = 1 To this.NodeTable.DataBodyRange.Rows.count
        
        Select Case UCase$(this.NodeTable.DataBodyRange.Item(row, 4))
        Case Pin
            xConstrained = True
            yConstrained = True
        Case Roller
            xConstrained = False
            yConstrained = True
        Case Else
            xConstrained = False
            yConstrained = False
        End Select
        
        With this
            .nodes.Add MakeNode2D( _
                .NodeTable.DataBodyRange.Item(row, 1), _
                MakePoint2D(.NodeTable.DataBodyRange.Item(row, 2), .NodeTable.DataBodyRange.Item(row, 3)), _
                xConstrained, _
                yConstrained _
                )
        End With
    Next row
End Sub

Private Sub GetMaterials()
    Dim row As Long
    For row = 1 To this.MaterialTable.DataBodyRange.Rows.count
        With this
            ' add array with cross-sectional area and Modulus of Elasticity to collection
            .Materials.Add .MaterialTable.DataBodyRange.Item(row, 2).value
        End With
    Next row
End Sub

Private Sub GetCrossSections()
    Dim row As Long
    For row = 1 To this.CrossSectionTable.DataBodyRange.Rows.count
        With this
            .CrossSections.Add .MaterialTable.DataBodyRange.Item(row, 2)
        End With
    Next row
End Sub

Private Sub GetMembers()

    Dim StartID As Long
    Dim EndID As Long
    Dim StartNode As Node2D
    Dim EndNode As Node2D
    Dim CrossSectionID As Long
    Dim MaterialID As Long
    Dim Area As Double
    Dim youngsMod As Double
    
    Dim row As Long
    For row = 1 To this.MemberTable.DataBodyRange.Rows.count
        With this
        
            StartID = .MemberTable.DataBodyRange.Item(row, 2)
            EndID = .MemberTable.DataBodyRange.Item(row, 3)
            Set StartNode = .nodes.Item(StartID)
            Set EndNode = .nodes.Item(EndID)
            CrossSectionID = .MemberTable.DataBodyRange.Item(row, 4)
            MaterialID = .MemberTable.DataBodyRange.Item(row, 5)
            Area = .CrossSections.Item(MaterialID)
            youngsMod = .Materials.Item(MaterialID)
            
            .Members.Add MakeBarElement2D( _
                .MemberTable.DataBodyRange.Item(row, 1), _
                StartNode, _
                EndNode, _
                Area, _
                youngsMod _
                )
                
        End With
    Next row
End Sub

Private Sub GetLoads()
    Dim nodeID As Long
    Dim xLoad As Double
    Dim yLoad As Double
    Dim currentNode As Node2D
    Dim row As Long
    For row = 1 To this.LoadTable.DataBodyRange.Rows.count
        With this
            nodeID = .LoadTable.DataBodyRange.Item(row, 1)
            xLoad = .LoadTable.DataBodyRange.Item(row, 2)
            yLoad = .LoadTable.DataBodyRange.Item(row, 3)
            Set currentNode = .nodes.Item(nodeID)
        End With
        
        currentNode.AddLoad MakeVector(xLoad, yLoad)
    Next row
End Sub
