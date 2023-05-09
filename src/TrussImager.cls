VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrussImager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Output")
Option Explicit

Private Type TTrussDrawer
    Target As Worksheet
    transform As AffineTransform
    nodeRadiusModifier As Double
    WindowMargin As Double
End Type

Private this As TTrussDrawer

Private Const NodeRadius As Double = 4
Private Const NodeDivisions As Long = 10
Private Const CompressionColor As Long = 16711680   ' Blue
Private Const TensionColor As Long = 255            ' Red
Private Const NodeColor As Long = 0                 ' Black
Private Const LoadArrowColor As Long = 36095        ' Orange
Private Const ReactionArrowColor As Long = 65280    ' Green
Private Const LoadArrowLength As Long = 20

Private Sub Class_Initialize()
    With this
        Set .transform = New AffineTransform
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set this.Target = Nothing
        Set this.transform = Nothing
    End With
End Sub

Public Property Set Target(ByRef sheet As Worksheet)
    Set this.Target = sheet
End Property

Public Property Set transform(ByRef Trans As AffineTransform)
    Set this.transform = Trans
End Property

Private Sub FitWindow(ByRef truss As SolutionStructure)

    Const marginRatio As Double = 0.75
    
    Dim box As Rectangle2D
    Set box = truss.BoundingRectangle(0)
    
    this.WindowMargin = WorksheetFunction.MIN(box.size.Height, box.size.Width) * marginRatio
    
    Set box = truss.BoundingRectangle(this.WindowMargin)
    
    Dim maxWidth As Double
    maxWidth = Application.Windows(1).UsableWidth
    
    Dim maxHeight As Double
    maxHeight = Application.Windows(1).UsableHeight

    Dim ScaleX As Double
    ScaleX = maxWidth / box.size.Width
    
    Dim ScaleY As Double
    ScaleY = maxHeight / box.size.Height
    
    Dim bestRatio As Double
    bestRatio = WorksheetFunction.MIN(ScaleX, ScaleY)
    
    this.nodeRadiusModifier = bestRatio

    With this.transform
        .ScaleX = bestRatio
        .ScaleY = -bestRatio
        .translateX = this.WindowMargin * bestRatio
        .translateY = (box.size.Height - this.WindowMargin) * bestRatio
    End With

End Sub

Public Sub Draw(ByRef truss As SolutionStructure)

    ClearShapes
    
    FitWindow truss
    DrawMembers truss
    DrawLoads truss
    DrawNodes truss

End Sub

Private Sub DrawMembers(ByRef truss As SolutionStructure)
    Dim currentSegment As Segment2D
    Set currentSegment = New Segment2D
    Dim bar As SolutionBarElement2D
    For Each bar In truss.Bars
        With currentSegment
            Set .StartPoint = bar.StartNode.OriginalPosition
            Set .EndPoint = bar.EndNode.OriginalPosition
        End With
        
        If IsEffectivelyZero(bar.InternalForceValue) Then
            DrawSegment currentSegment, CStr(bar.ID), RGB(0, 0, 0)
        ElseIf bar.InternalForceValue < 0 Then
            DrawSegment currentSegment, CStr(bar.ID), CompressionColor
        ElseIf bar.InternalForceValue > 0 Then
            DrawSegment currentSegment, CStr(bar.ID), TensionColor
        End If
        
    Next
End Sub

Private Sub DrawNodes(ByRef truss As SolutionStructure)
    Dim Node As SolutionNode2D
    For Each Node In truss.Nodes
        DrawNode Node
    Next Node
End Sub

Private Sub DrawLoads(ByRef truss As SolutionStructure)
    Dim Node As SolutionNode2D
    For Each Node In truss.Nodes
        If Node.IsLoaded Then
            DrawNodalLoads Node
        End If
    Next Node
End Sub

Private Sub DrawNodalLoads(ByRef Node As SolutionNode2D)
    Const bgArrowNudge As Long = 3
    
    Dim transNode As Point2D
    Set transNode = this.transform.ApplyToPoint(Node.OriginalPosition)
    
    Dim Load As Vector2D
    Set Load = Node.NetLoad
    
    If Not IsEffectivelyZero(Load.u) Then
        With this.Target.Shapes.AddLine(transNode.x, transNode.y, transNode.x + (LoadArrowLength + bgArrowNudge) * Math.Sgn(Load.u), transNode.y)
            .Name = Node.ID & "-xbg"
            .Line.Weight = 4
            .Line.ForeColor.RGB = NodeColor
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
        With this.Target.Shapes.AddLine(transNode.x, transNode.y, transNode.x + LoadArrowLength * Math.Sgn(Load.u), transNode.y)
            .Name = Node.ID & "-x"
            .Line.Weight = 2
            .Line.ForeColor.RGB = LoadArrowColor
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
    End If
    
    If Not IsEffectivelyZero(Load.v) Then
        With this.Target.Shapes.AddLine(transNode.x, transNode.y, transNode.x, transNode.y - (LoadArrowLength + bgArrowNudge) * Math.Sgn(Load.v))
            .Name = Node.ID & "-ybg"
            .Line.Weight = 4
            .Line.ForeColor.RGB = NodeColor
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
        With this.Target.Shapes.AddLine(transNode.x, transNode.y, transNode.x, transNode.y - LoadArrowLength * Math.Sgn(Load.v))
            .Name = Node.ID & "-y"
            .Line.Weight = 2
            .Line.ForeColor.RGB = LoadArrowColor
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
    End If
End Sub

Private Sub ClearShapes()
    Dim shape As shape
    For Each shape In this.Target.Shapes
        shape.Delete
    Next shape
End Sub

Private Sub DrawSegment(ByRef segment As Segment2D, ByRef BarID As String, ByRef Color As Long)

    Dim transSegment As Segment2D
    Set transSegment = this.transform.ApplyToSegment(segment)
    
    With this.Target.Shapes.AddLine(transSegment.StartPoint.x, transSegment.StartPoint.y, transSegment.EndPoint.x, transSegment.EndPoint.y)
        .Name = BarID
        .Line.Weight = 2
        .Line.ForeColor.RGB = Color
    End With
End Sub

Private Sub DrawNode(ByRef Node As SolutionNode2D)

    Dim transNode As Polygon2D
    Set transNode = this.transform.ApplyToPolygon(MakeCircle(NodeRadius / this.nodeRadiusModifier, Node.OriginalPosition).ToPolygon(NodeDivisions))
    
    Dim circlePoints() As Single
    ReDim circlePoints(transNode.Vertices.Count, 0 To 1)
    
    Dim currentPoint As Point2D
    Dim i As Long
    For i = 0 To transNode.Vertices.Count - 1
        Set currentPoint = transNode.Vertices.Item(i + 1)
        circlePoints(i, 0) = CSng(currentPoint.x)
        circlePoints(i, 1) = CSng(currentPoint.y)
    Next i
    
    Set currentPoint = transNode.Vertices.Item(1)
    circlePoints(transNode.Vertices.Count, 0) = CSng(currentPoint.x)
    circlePoints(transNode.Vertices.Count, 1) = CSng(currentPoint.y)
    
    With this.Target.Shapes.AddPolyline(circlePoints)
        .Name = CStr(Node.ID)
        .Fill.BackColor.RGB = NodeColor
        .Fill.ForeColor.RGB = NodeColor
    End With
    
End Sub
