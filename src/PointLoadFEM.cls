VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointLoadFEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.FixedEndForces")
'@Description("Fixed end forces for beam element with point load.")
Option Explicit
Implements IFixedEndForce

Private Type tPointLoadFEM
    BeamLoad As Double
    StartToLoadDistance As Double
    EndToLoadDistance As Double
    BeamLength As Double
End Type

Private this As tPointLoadFEM
Private Const CLASS_NAME As String = "PointLoadFEM"

Public Property Get BeamLoad() As Double
    BeamLoad = this.BeamLoad
End Property

Public Property Let BeamLoad(ByVal value As Double)
    this.BeamLoad = value
End Property

Public Property Get BeamLength() As Double
    BeamLength = this.BeamLength
End Property

Public Property Let BeamLength(ByVal value As Double)

    If value <= 0 Then StructuralModelErrors.ThrowBadElementLengthError CLASS_NAME
    
    this.BeamLength = value
    
End Property

Public Property Get DistanceToLoad() As Double
    DistanceToLoad = this.StartToLoadDistance
End Property

Public Property Let DistanceToLoad(ByVal value As Double)

    If value < 0 Or value > this.BeamLength Then StructuralModelErrors.ThrowBadLoadLocationError CLASS_NAME
    
    this.StartToLoadDistance = value
    this.EndToLoadDistance = this.BeamLength - this.StartToLoadDistance
    
End Property

Public Function FixedEndMoment(ByVal location As FixedEndForceLocation) As Double
    
    Select Case location
    Case FixedEndForceLocation.Left
        FixedEndMoment = LeftNodeEndMoment
    Case FixedEndForceLocation.Right
        FixedEndMoment = RightNodeEndMoment
    End Select
    
End Function

Public Function FixedEndReaction(ByVal location As FixedEndForceLocation) As Double
    
    Select Case location
    Case FixedEndForceLocation.Left
        FixedEndReaction = LeftNodeEndReaction
    Case FixedEndForceLocation.Right
        FixedEndReaction = RightNodeEndReaction
    End Select
    
End Function

Private Function LeftNodeEndMoment() As Double
    
    With this
        Dim numerator As Double
        numerator = .BeamLoad * .StartToLoadDistance * .EndToLoadDistance ^ 2
    End With

    LeftNodeEndMoment = numerator / this.BeamLength ^ 2
    
End Function

Private Function RightNodeEndMoment() As Double
    
    With this
        Dim numerator As Double
        numerator = .BeamLoad * .StartToLoadDistance ^ 2 * .EndToLoadDistance
    End With

    RightNodeEndMoment = numerator / this.BeamLength ^ 2
    
End Function

Private Function LeftNodeEndReaction() As Double
    
    With this
        Dim numerator As Double
        numerator = .BeamLoad * .EndToLoadDistance ^ 2 * (3 * .StartToLoadDistance + .EndToLoadDistance)
    End With

    LeftNodeEndReaction = numerator / this.BeamLength ^ 3
    
End Function

Private Function RightNodeEndReaction() As Double
    
    With this
        Dim numerator As Double
        numerator = .BeamLoad * .StartToLoadDistance ^ 2 * (.StartToLoadDistance + 3 * .EndToLoadDistance)
    End With

    RightNodeEndReaction = numerator / this.BeamLength ^ 3
    
End Function


' IFixedEndForce Implementation

Private Function IFixedEndForce_FixedEndMoment(ByVal location As FixedEndForceLocation) As Double
    IFixedEndForce_FixedEndMoment = Me.FixedEndMoment(location)
End Function

Private Function IFixedEndForce_FixedEndReaction(ByVal location As FixedEndForceLocation) As Double
    IFixedEndForce_FixedEndReaction = Me.FixedEndReaction(location)
End Function
