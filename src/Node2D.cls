VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.Structure Model")
Option Explicit

Private Type TNode
    nodeID As Long
    Position As Point2D
    Loads As Collection
    xConstrained As Boolean
    yConstrained As Boolean
End Type

Private this As TNode

Private Sub Class_Initialize()
    With this
        Set .Position = New Point2D
        Set .Loads = New Collection
        .xConstrained = False   ' default is that node can freely translate in x and y directions
        .yConstrained = False
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Position = Nothing
        Set .Loads = Nothing
    End With
End Sub

Public Property Get ID() As Long
    ID = this.nodeID
End Property

Public Property Let ID(ByRef nodeID As Long)
    this.nodeID = nodeID
End Property

Public Property Get Position() As Point2D
    Set Position = this.Position
End Property

Public Property Set Position(ByRef nodeCoordinates As Point2D)
    Set this.Position = nodeCoordinates
End Property

Public Property Get Loads() As Collection
    Set Loads = this.Loads
End Property

Public Property Set Loads(ByRef nodeLoads As Collection)
    Set this.Loads = nodeLoads
End Property

Public Property Get LoadCount() As Long
    LoadCount = this.Loads.count
End Property

Public Property Get xConstrained() As Boolean
    xConstrained = this.xConstrained
End Property

Public Property Let xConstrained(ByRef constrained As Boolean)
    this.xConstrained = constrained
End Property

Public Property Get yConstrained() As Boolean
    yConstrained = this.yConstrained
End Property

Public Property Let yConstrained(ByRef constrained As Boolean)
    this.yConstrained = constrained
End Property

Public Property Get NetLoad() As Vector2D
    
    Dim resultant As Vector2D
    Set resultant = New Vector2D
    resultant.u = 0
    resultant.v = 0
        
    If Me.LoadCount > 0 Then
        Dim index As Long
        For index = 1 To Me.LoadCount
            Set resultant = resultant.Add(this.Loads.Item(index))
        Next index
    End If
    
    Set NetLoad = resultant
    
End Property

Public Sub AddLoad(ByRef loadVector As Vector2D)
    this.Loads.Add loadVector
End Sub

Public Function Equals(ByRef other As Node2D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Node2D Then ' other is not a Node2D
        Equals = False
    Else
        Equals = Me.Position.Equals(other.Position)
    End If

End Function
