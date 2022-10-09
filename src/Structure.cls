VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Public Enum StructureErrors
    Unstable = 800 + vbObjectError
End Enum

Private Type TStructure
    Bars As Collection
    Nodes As Collection
    DOF_Dict As Scripting.Dictionary
    SystemMatrix As Matrix
    SystemVector As Vector
    GlobalDisplacements As Vector
End Type

Private Const DOFsPerNode As Long = 2

Private this As TStructure

Private Sub Class_Initialize()
    With this
        Set .Bars = New Collection
        Set .Nodes = New Collection
        Set .DOF_Dict = New Scripting.Dictionary
        Set .SystemMatrix = New Matrix
        Set .SystemVector = New Vector
        Set .GlobalDisplacements = New Vector
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Bars = Nothing
        Set .Nodes = Nothing
        Set .DOF_Dict = Nothing
        Set .SystemMatrix = Nothing
        Set .SystemVector = Nothing
        Set .GlobalDisplacements = Nothing
    End With
End Sub

Public Property Get Bars() As Collection
    Set Bars = this.Bars
End Property

Public Property Set Bars(ByRef barCollection As Collection)
    Set this.Bars = barCollection
End Property

Public Property Get Nodes() As Collection
    Set Nodes = this.Nodes
End Property

Public Property Set Nodes(ByRef nodeCollection As Collection)
    Set this.Nodes = nodeCollection
End Property

Public Property Get NodeCount() As Long
    NodeCount = this.Nodes.count
End Property

Public Property Get BarCount() As Long
    BarCount = this.Bars.count
End Property

Public Property Get LoadCount() As Long

    Dim count As Long
    count = 0
    
    Dim Node As Node2D
    
    Dim index As Long
    For index = 1 To Me.NodeCount
        Set Node = this.Nodes.Item(index)
        count = count + Node.LoadCount
    Next index
    
    LoadCount = count

End Property

Public Property Get ReactionCount() As Long
    Dim count As Long
    count = 0
    
    Dim Node As Node2D
    Dim index As Long
    For index = 1 To Me.NodeCount
        Set Node = this.Nodes.Item(index)
        If Node.xConstrained Then count = count + 1
        If Node.yConstrained Then count = count + 1
    Next index
    
    ReactionCount = count
End Property

Public Function IsStable() As Boolean
    IsStable = BarCount + ReactionCount >= 2 * NodeCount
End Function

Public Function IsDeterminate() As Boolean
    IsDeterminate = BarCount + ReactionCount = 2 * NodeCount
End Function

Public Function DegreeOfIndeterminancy() As Long
    DegreeOfIndeterminancy = BarCount + ReactionCount - 2 * NodeCount
End Function

Public Function Solve() As SolutionStructure

    If Not IsStable Then
        Err.Raise StructureErrors.Unstable, "Structer.Solve", _
            "Truss is unstable. Check support conditions."
    End If
    
    AssignDegreesOfFreedom
    SolveSystemOfEquations
    Set Solve = MakeStructureSolution
End Function

Private Sub AssignDegreesOfFreedom()
    Set this.DOF_Dict = New Scripting.Dictionary
    Dim index As Long
    For index = 1 To this.Nodes.count
        this.DOF_Dict.Add Key:=this.Nodes.Item(index).ID, Item:=Array(2 * (index - 1), 2 * (index - 1) + 1)
    Next index
End Sub

Private Sub SolveSystemOfEquations()
    Dim size As Long
    size = Me.NodeCount * DOFsPerNode
    
    AssembleSystemMatrix size
    AssembleSystemVector size
    ApplyExternalConstraints
    
    Dim solver As Cholesky
    Set solver = New Cholesky
    Set this.GlobalDisplacements = solver.Solve(this.SystemMatrix, this.SystemVector)
End Sub

Private Function MakeStructureSolution() As SolutionStructure
    
    Dim Nodes As Collection
    Set Nodes = New Collection
    Dim index As Long
    For index = 1 To this.Nodes.count
        Nodes.Add NodeToSolution(this.Nodes.Item(index))
    Next index
    
    ' Create nodes dictionary to aid in creating SolutionBarElement2D collection
    Dim NodesDict As Scripting.Dictionary
    Set NodesDict = New Scripting.Dictionary
    For index = 1 To Nodes.count
        NodesDict.Add Key:=Nodes.Item(index).ID, Item:=Nodes.Item(index)
    Next index
    
    Dim Bars As Collection
    Set Bars = New Collection
    Dim currentBar As BarElement2D
    For index = 1 To this.Bars.count
        Set currentBar = this.Bars.Item(index)
        Bars.Add MakeSolutionBar(currentBar, _
                                 NodesDict(currentBar.StartNode.ID), _
                                 NodesDict(currentBar.EndNode.ID))
    Next index
    
    Dim solution As SolutionStructure
    Set solution = New SolutionStructure
    With solution
        Set .Bars = Bars
        Set .Nodes = Nodes
    End With
    
    Set MakeStructureSolution = solution
    
End Function

Private Function NodeToSolution(ByRef structureNode As Node2D) As SolutionNode2D
    Dim dofX As Long
    dofX = this.DOF_Dict.Item(structureNode.ID)(0)  ' x-direction translational degree-of-freedom
    
    Dim dofY As Long
    dofY = this.DOF_Dict.Item(structureNode.ID)(1)  ' y-direction translational degree-of-freedom
    
    Dim displacement As Vector2D
    Set displacement = New Vector2D
    With displacement
        .u = this.GlobalDisplacements.ValueAt(dofX)
        .v = this.GlobalDisplacements.ValueAt(dofY)
    End With
    
    Dim solutionNode As SolutionNode2D
    Set solutionNode = New SolutionNode2D
    solutionNode.SetSolutionNode originalNode:=structureNode, displacedPosition:=displacement
    
    Set NodeToSolution = solutionNode
    
End Function

Private Sub AssembleSystemMatrix(ByRef size As Long)

    Dim mat As Matrix
    Set mat = New Matrix
    mat.SetSize size, size
    
    Dim barMatrix As Matrix
    Set barMatrix = New Matrix
    
    Dim barDofs As Collection
    Set barDofs = New Collection
    
    Dim rowIndex As Long
    Dim colIndex As Long
    
    Dim barIndex As Long
    For barIndex = 1 To this.Bars.count
        Set barMatrix = this.Bars.Item(barIndex).GlobalStiffnessMatrix
        Set barDofs = BarDegreesOfFreedom(this.Bars.Item(barIndex))
        For rowIndex = 1 To barDofs.count
            For colIndex = 1 To barDofs.count
                Set mat = mat.AddTo(barDofs.Item(rowIndex), _
                                    barDofs.Item(colIndex), _
                                    barMatrix.ValueAt(rowIndex - 1, colIndex - 1))
            Next colIndex
        Next rowIndex
        
    Next barIndex
    
    Set this.SystemMatrix = mat

End Sub

Private Function BarDegreesOfFreedom(ByRef bar As BarElement2D) As Collection
    Dim dofs As Collection
    Set dofs = New Collection
    Dim index As Long
    For index = 1 To this.DOF_Dict.count
        If index = bar.StartNode.ID Or index = bar.EndNode.ID Then
            dofs.Add this.DOF_Dict.Item(index)(0)
            dofs.Add this.DOF_Dict.Item(index)(1)
        End If
    Next index
    
    Set BarDegreesOfFreedom = dofs
End Function

Private Sub AssembleSystemVector(ByRef size As Long)

    Dim vec As Vector
    Set vec = New Vector
    Set vec = vec.SetLength(size)
    
    Dim net_Load As Vector2D
    Set net_Load = New Vector2D
    Dim nodeIndex As Long
    For nodeIndex = 1 To this.Nodes.count
        Set net_Load = this.Nodes.Item(nodeIndex).NetLoad
        Set vec = vec.AddTo(CLng(this.DOF_Dict.Item(nodeIndex)(0)), net_Load.u)
        Set vec = vec.AddTo(CLng(this.DOF_Dict.Item(nodeIndex)(1)), net_Load.v)
    Next nodeIndex
    
    Set this.SystemVector = vec

End Sub

Private Sub ApplyExternalConstraints()

    Dim dofX As Long
    Dim dofY As Long
    Dim nodeIndex As Long
    Dim currentNode As Node2D
    For nodeIndex = 1 To this.Nodes.count
        Set currentNode = this.Nodes.Item(nodeIndex)
        dofX = this.DOF_Dict(currentNode.ID)(0)
        dofY = this.DOF_Dict(currentNode.ID)(1)
        
        If currentNode.xConstrained Then
            this.SystemMatrix.SetIdentityRow rowIndex:=dofX
            this.SystemMatrix.SetIdentityColumn columnIndex:=dofX
            this.SystemVector.ValueAt(dofX) = 0
        End If
        
        If currentNode.yConstrained Then
            this.SystemMatrix.SetIdentityRow rowIndex:=dofY
            this.SystemMatrix.SetIdentityColumn columnIndex:=dofY
            this.SystemVector.ValueAt(dofY) = 0
        End If
    Next nodeIndex

End Sub
