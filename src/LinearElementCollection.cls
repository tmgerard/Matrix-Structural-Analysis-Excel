VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearElementCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TLinearElementCollection
    Elements As Collection
End Type

Private this As TLinearElementCollection

Private Sub Class_Initialize()
    Set this.Elements = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Elements = Nothing
End Sub

Public Sub Add(ByRef Element As ILinearElement)
    If this.Elements.Count = 0 Then
        Element.ID = 0
        this.Elements.Add Element
    Else
        Dim lastElement As ILinearElement
        Set lastElement = this.Elements(this.Elements.Count)
        Element.ID = lastElement.ID + 1
        this.Elements.Add Element
    End If
End Sub

Public Function Count() As Long
    Count = this.Elements.Count
End Function

Public Function Item(ByRef elementIndex As Long) As ILinearElement
    Set Item = this.Elements.Item(elementIndex)
End Function

Public Sub Remove(ByRef elementIndex As Long)
    Dim toRemove As ILinearElement
    Set toRemove = this.Elements(elementIndex)
    
    Dim removeID As Long
    removeID = toRemove.ID
    
    this.Elements.Remove elementIndex
    
    Dim Element As ILinearElement
    For Each Element In this.Elements
        If Element.ID > removeID Then Element.ID = Element.ID - 1
    Next Element
End Sub
