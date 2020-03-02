VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.Geometry.Interface")
'@Interface ICoordinate
'@IgnoreModule UseMeaningfulName
Option Explicit

Public Property Get X() As Double
End Property

Public Property Get Y() As Double
End Property

Public Property Get Z() As Double
End Property

Public Function Distance(ByVal other As ICoordinate) As Double
End Function
