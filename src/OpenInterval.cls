VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Geometry")
Option Explicit

Private Type TOpenInterval
    Start As Double
    End As Double
End Type

Private this As TOpenInterval

Public Enum OpenIntervalErrors
    BadInterval = 300 + vbObjectError
End Enum

Public Sub SetInterval(ByRef IntStart As Double, ByRef IntEnd As Double)
    If IntStart > IntEnd Then
        Err.Raise Number:=OpenIntervalErrors.BadInterval, _
                  source:="OpenInterval.SetInterval", _
                  Description:="Interval start value must be less than the end value."
    End If
    this.Start = IntStart
    this.End = IntEnd
End Sub

Public Property Get IntStart() As Double
    IntStart = this.Start
End Property

Public Property Get IntEnd() As Double
    IntEnd = this.End
End Property

Public Function Length() As Double
    Length = this.End - this.Start
End Function

Public Function Contains(ByRef value As Double) As Boolean
    Contains = this.Start < value And value < this.End
End Function

Public Function OverlapsInterval(ByRef other As OpenInterval) As Boolean
    If Doubles.Equal(this.Start, other.IntStart) And _
       Doubles.Equal(this.End, other.IntEnd) Then
        OverlapsInterval = True
        Exit Function
    End If
    
    OverlapsInterval = Me.Contains(other.IntStart) Or _
                       Me.Contains(other.IntStart) Or _
                       Me.Contains(this.Start) Or _
                       Me.Contains(this.End)
End Function

Public Function ComputeOverlapWith(ByRef other As OpenInterval) As OpenInterval
    If Not Me.OverlapsInterval(other) Then
        Set ComputeOverlapWith = Nothing
        Exit Function
    End If
    
    Dim interval As OpenInterval
    Set interval = New OpenInterval
    interval.SetInterval WorksheetFunction.MAX(this.Start, other.IntStart), _
                         WorksheetFunction.MIN(this.End, other.IntEnd)
    
    Set ComputeOverlapWith = interval
End Function
