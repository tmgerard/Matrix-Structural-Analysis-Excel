VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolutionOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Output")
Option Explicit

Private Type TSolutionOutput
    solution As SolutionStructure
    outputSheet As Worksheet
End Type

Private this As TSolutionOutput

Private Const DecimalFormat As String = "#0.000"

Private Sub Class_Terminate()
    With this
        Set .solution = Nothing
        Set .outputSheet = Nothing
    End With
End Sub

Public Sub WriteToWorksheet(ByRef solution As SolutionStructure, ByRef outputSheet As Worksheet)
    
    With outputSheet
        .Cells.Clear
        .Cells(1, 1) = "Truss Solution"
        .Cells(2, 1) = solution.Bars.count & " Members"
        .Cells(3, 1) = solution.nodes.count & " Nodes"
    End With
    
    With this
        Set .solution = solution
        Set .outputSheet = outputSheet
    End With
    
    Dim row As Long
    row = 5
    
    this.outputSheet.Cells(row, 1) = "----------[Node Information]----------"
    row = row + 1
    
    NodeToSheet row
    
    this.outputSheet.Cells(row, 1) = "----------[Element Information]----------"
    row = row + 1
    BarsToSheet row
    
End Sub

Private Sub NodeToSheet(ByRef row As Long)
    Dim currentNode As SolutionNode2D
    
    Dim I As Long
    For I = 1 To this.solution.nodes.count
        Set currentNode = this.solution.nodes.Item(I)
        
        With currentNode
            this.outputSheet.Cells(row, 1).Value2 = "Node " & .ID
            row = row + 1
            
            this.outputSheet.Cells(row, 1) = vbTab & "Original Pos: (" & Format(.OriginalPosition.x, DecimalFormat) _
                & ", " & Format(.OriginalPosition.y, DecimalFormat) & ")"
            row = row + 1
            
            this.outputSheet.Cells(row, 1) = vbTab & "Displaced Pos: (" & Format(.displacedPosition.x, DecimalFormat) _
                & ", " & Format(.displacedPosition.y, DecimalFormat) & ")"
            row = row + 1
            
            If .IsConstrained Then
                Dim reaction As Vector2D
                Set reaction = this.solution.NodeReaction(currentNode)
                this.outputSheet.Cells(row, 1) = vbTab & "Reaction: Fx = " & Format(reaction.u, DecimalFormat) & _
                    " Fy = " & Format(reaction.v, DecimalFormat)
                row = row + 1
            End If
        End With
        
        row = row + 1
    Next I
    
End Sub

Private Sub BarsToSheet(ByRef row As Long)
    Dim currentBar As SolutionBarElement2D
    Dim nodeString As String
    Dim loadType As String
    Dim elongation As String
    Dim strain As String
    Dim stress As String
    Dim barForce As String
    
    Dim I As Long
    For I = 1 To this.solution.Bars.count
        Set currentBar = this.solution.Bars.Item(I)
        
        nodeString = "Node " & currentBar.StartNode.ID & " to Node " & currentBar.EndNode.ID
        
        If currentBar.stress >= 0 Then
            loadType = "TENSION"
        Else
            loadType = "COMPRESSION"
        End If
        
        elongation = Format$(currentBar.elongation, DecimalFormat)
        strain = Format$(currentBar.strain, "Scientific")
        stress = Format$(currentBar.stress, DecimalFormat)
        barForce = Format$(currentBar.stress * currentBar.Area, DecimalFormat)
        
        this.outputSheet.Cells(row, 1).Value2 = "Bar " & currentBar.ID & " - " & nodeString & " : " & loadType
        row = row + 1
        this.outputSheet.Cells(row, 1).Value2 = vbTab & "Elongation = " & elongation
        row = row + 1
        this.outputSheet.Cells(row, 1).Value2 = vbTab & "Strain = " & strain
        row = row + 1
        this.outputSheet.Cells(row, 1).Value2 = vbTab & "Stress = " & stress
        row = row + 1
        this.outputSheet.Cells(row, 1).Value2 = vbTab & "Axial Force = " & barForce
        
        row = row + 2
    Next I
End Sub
