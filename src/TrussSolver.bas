Attribute VB_Name = "TrussSolver"
'@Folder("StructuralAnalysis")
Option Explicit

Public Sub Main()

    On Error GoTo ErrorHandler

    Dim builder As StructureBuilder
    Set builder = New StructureBuilder
    
    Dim truss As Structure
    Set truss = builder.Build
    
    Set builder = Nothing
    
    Dim startTime As Single
    Dim endTime As Single
    Dim analysisTime As Single
    startTime = Timer
    
    Application.StatusBar = "Analyzing truss"
    Dim solution As SolutionStructure
    Set solution = truss.Solve
    
    endTime = Timer
    analysisTime = endTime - startTime
    
    Dim solutionReporter As SolutionOutput
    Set solutionReporter = New SolutionOutput
    
    Application.StatusBar = "Creating truss analysis report"
    solutionReporter.WriteToWorksheet solution, TrussOutput
    
    Dim Trans As AffineTransform
    Set Trans = New AffineTransform
    Trans.ScaleX = 0.5
    Trans.ScaleY = -0.5
    Trans.translateY = 300
    Trans.translateX = 15
    
    Dim drawer As TrussImager
    Set drawer = New TrussImager
    With drawer
        Set .Target = TrussDrawing
        Set .transform = Trans
    End With
    
    Application.StatusBar = "Drawing truss"
    drawer.Draw solution
    
    Application.StatusBar = "Analysis Completed in " & Format(analysisTime, "#0.0000") & " Seconds"
    TrussOutput.Activate

SubExit:
    Set builder = Nothing
    Set truss = Nothing
    Set solution = Nothing
Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    Resume SubExit

End Sub
