VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.Node")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements INode

Private Type TNode3D
    nodeDescription As String
    Coordinate As Coordinate3d
    activeDOF() As Boolean
End Type

Private this As TNode3D
Private Const CLASS_NAME As String = "Node3D"
Private Const TotalDegreesOfFreedom As Integer = 6

Private Sub Class_Initialize()

    Set this.Coordinate = New Coordinate2d
    
    ReDim this.activeDOF(TotalDegreesOfFreedom - 1)
    
    ' initialize all nodal dofs as active
    Dim dof As Integer
    For dof = LBound(this.activeDOF) To UBound(this.activeDOF)
        this.activeDOF(dof) = True
    Next dof
    
End Sub

Private Sub Class_Terminate()
    Set this.Coordinate = Nothing
End Sub

Public Property Get Coordinate() As ICoordinate
    Set Coordinate = this.Coordinate
End Property

Public Property Set Coordinate(ByRef value As Coordinate3d)
    Set this.Coordinate = value
End Property

Public Property Get Description() As String
    Description = this.nodeDescription
End Property

Public Property Let Description(ByVal value As String)
    this.nodeDescription = value
End Property

Public Function DegreesOfFreedom() As Integer
    DegreesOfFreedom = TotalDegreesOfFreedom
End Function

'INode Implementation

Private Property Get INode_Coordinate() As ICoordinate
    Set INode_Coordinate = this.Coordinate
End Property
