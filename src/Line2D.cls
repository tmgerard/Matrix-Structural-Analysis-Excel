VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TLine2D
    Base As Point2D
    direction As Vector2D
End Type

Private this As TLine2D

Private Sub Class_Initialize()
    With this
        Set .Base = New Point2D
        Set .direction = New Vector2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Base = Nothing
        Set .direction = Nothing
    End With
End Sub

Public Property Get Base() As Point2D
    Set Base = this.Base
End Property

Public Property Set Base(ByRef point As Point2D)
    Set this.Base = point
End Property

Public Property Get direction() As Vector2D
    Set direction = this.direction
End Property

Public Property Set direction(ByRef Vector As Vector2D)
    Set this.direction = Vector
End Property

'@Description "Checks if Line2D objects are parallel."
Public Function IsParallelTo(ByRef other As Line2D) As Boolean
Attribute IsParallelTo.VB_Description = "Checks if Line2D objects are parallel."
    IsParallelTo = this.direction.IsParallelTo(other.direction)
End Function

'@Description "Checks if Line2D objects are perpendicular."
Public Function IsPerpendicularTo(ByRef other As Line2D) As Boolean
Attribute IsPerpendicularTo.VB_Description = "Checks if Line2D objects are perpendicular."
    IsPerpendicularTo = this.direction.IsPerpendicularTo(other.direction)
End Function

'@Description "Creates Line2D object through a given point that is perpendicular to the existing Line2D object."
Public Function PerpendicularThrough(ByRef point As Point2D) As Line2D
Attribute PerpendicularThrough.VB_Description = "Creates Line2D object through a given point that is perpendicular to the existing Line2D object."
    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        .Base = point
        .direction = this.direction.Perpendicular
    End With
    
    Set PerpendicularThrough = newLine
End Function

'@Description "Creates Line2D object through a given point that is parallel to the existing Line2D object."
Public Function ParallelThrough(ByRef point As Point2D) As Line2D
Attribute ParallelThrough.VB_Description = "Creates Line2D object through a given point that is parallel to the existing Line2D object."
    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        .Base = point
        .direction = this.direction
    End With
    
    Set ParallelThrough = newLine
End Function

'@Description "Returns Point2D object representing the intersection between two Line2D objects."
Public Function IntersectionWith(ByRef other As Line2D) As Point2D
Attribute IntersectionWith.VB_Description = "Returns Point2D object representing the intersection between two Line2D objects."
    If Me.IsParallelTo(other) Then
        Set IntersectionWith = Nothing
        Exit Function
    End If
    
    Dim d1 As Vector2D
    Set d1 = this.direction
    
    Dim d2 As Vector2D
    Set d2 = other.direction
    
    Dim Cross As Double
    Cross = d1.Cross(d2)
    
    Dim delta As Vector2D
    Set delta = Vector2DFactory.MakeVectorBetween(this.Base, other.Base)
    
    Dim t1 As Double
    t1 = (delta.u * d2.v - delta.v * d2.u) / Cross
    
    Set IntersectionWith = this.Base.Displaced(d1, t1)
End Function
