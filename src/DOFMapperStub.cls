VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DOFMapperStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Model.Structure Model")
' Stub for testing element connectivity manager
Option Explicit
Implements IDOFMapper

Private map As Matrix
Private Const Dofs As Long = 2

Private Sub Class_Initialize()
    Set map = New Matrix
    Set map = map.SetSize(2, 4)
    map.ValueAt(0, 1) = 1
    map.ValueAt(1, 1) = 2
    map.ValueAt(0, 3) = 3
    map.ValueAt(1, 3) = 4
End Sub

Private Sub Class_Terminate()
    Set map = Nothing
End Sub

Private Property Get IDOFMapper_DofMap() As Matrix
    Set IDOFMapper_DofMap = map
End Property

Private Property Get IDOFMapper_DofsPerNode() As Integer
    IDOFMapper_DofsPerNode = Dofs
End Property

Private Property Get IDOFMapper_GetDofEq(dofIndex As Long, nodeIndex As Long) As Long
    IDOFMapper_GetDofEq = map.ValueAt(dofIndex, nodeIndex)
End Property
