VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementConnectivityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TElementConnectivityManager
    map As IDOFMapper
    Elements As Collection
End Type

Private this As TElementConnectivityManager

Public Property Set DofMap(ByRef map As IDOFMapper)
    Set this.map = map
End Property

Public Property Set Elements(ByRef elementCollection As Collection)
    Set this.Elements = elementCollection
End Property

Public Function GetConnectivityArray(ByRef elementID As Long) As Long()
    
    Dim elem As IElementConnectivity
    Set elem = this.Elements.Item(elementID)
    
    Dim eNodes As Collection
    Set eNodes = elem.ElementNodes
    
    Dim result() As Long
    ReDim result(eNodes.count)
    
    Dim dofIndex As Long
    Dim eNode As Variant
    Dim elementDof As Long
    For Each eNode In eNodes
        For dofIndex = 0 To this.map.DofsPerNode
            elementDof = (eNode.ID - 1) * (this.map.DofsPerNode - 1) + dofIndex
            result(elementDof) = this.map.GetDofEq(dofIndex, eNode.ID - 1)
        Next dofIndex
    Next eNode
    
    
End Function
