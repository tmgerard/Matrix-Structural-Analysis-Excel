Attribute VB_Name = "SolutionBarElement2DFactory"
'@Folder("StructuralAnalysis.Model.Structure Solution")
Option Explicit

Public Function MakeSolutionBar(ByRef OriginalBar As BarElement2D, _
                                ByRef StartNode As SolutionNode2D, _
                                ByRef EndNode As SolutionNode2D) As SolutionBarElement2D

    Dim solutionBar As SolutionBarElement2D
    Set solutionBar = New SolutionBarElement2D
    solutionBar.SetSolutionBar OriginalBar:=OriginalBar, StartNode:=StartNode, EndNode:=EndNode
    
    Set MakeSolutionBar = solutionBar

End Function
