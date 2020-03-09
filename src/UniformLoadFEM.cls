VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UniformLoadFEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.FixedEndForces")
'@Description("Fixed end forces for uniformly loaded beam element.")
Option Explicit

Private Type tUniformLoadFEM
    BeamLoad As Double
    BeamLength As Double
End Type

Private this As tUniformLoadFEM
Private Const CLASS_NAME As String = "UniformLoadFEM"

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

    If value <= 0 Then StructuralModelErrors.ThrowBadElementLengthError ErrorSource:=CLASS_NAME
    
    this.BeamLength = value
    
End Property

Public Function FixedEndMoment(ByVal location As FixedEndForceLocation) As Double
    Select Case location
    
    Case FixedEndForceLocation.Left
        FixedEndMoment = EndMoment
    Case FixedEndForceLocation.Right
        FixedEndMoment = -EndMoment
    End Select
    
End Function

Public Function FixedEndReaction(ByVal location As FixedEndForceLocation) As Double
    FixedEndReaction = EndReaction
End Function

Private Function EndMoment() As Double
    EndMoment = (this.BeamLoad * this.BeamLength ^ 2) / 12
End Function

Private Function EndReaction() As Double
    EndReaction = (this.BeamLoad * this.BeamLength) / 2
End Function
