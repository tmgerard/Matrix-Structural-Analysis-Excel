VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlaneTrussDOFMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TPlaneTrussDOFMapper
    dofMat As Matrix
End Type

Private this As TPlaneTrussDOFMapper
Private Const MaxDOFsPerNode As Integer = 2

Public Property Get GetDofEq(ByRef elementIndex As Long, ByRef nodeIndex As Long) As Long
    GetDofEq = this.dofMat.ValueAt(elementIndex, nodeIndex)
End Property

Public Property Get DofMap() As Matrix
    Set DofMap = this.dofMat
End Property

Public Sub MapDofs(ByRef nodes As Collection)
    Set this.dofMat = New Matrix
    this.dofMat.SetSize MaxDOFsPerNode, nodes.count
    
    Dim currentNode As Node2D
    Dim eqNumber As Long
    eqNumber = 1
    Dim nodeIndex As Integer
    For nodeIndex = 1 To nodes.count
        Set currentNode = nodes.Item(nodeIndex)
        
        If Not currentNode.xConstrained Then
            this.dofMat.ValueAt(0, nodeIndex - 1) = eqNumber
            eqNumber = eqNumber + 1
        End If
        
        If Not currentNode.yConstrained Then
            this.dofMat.ValueAt(1, nodeIndex - 1) = eqNumber
            eqNumber = eqNumber + 1
        End If
    Next nodeIndex
End Sub


