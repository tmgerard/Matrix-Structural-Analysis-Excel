VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStructuralCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.StructuralModel.StructuralCrossSections.Interface")
'@Interface IStructuralCrossSection
'@IgnoreModule UseMeaningfulName
Option Explicit

'@Description("Cross sectional area.")
Public Function Area() As Double
End Function

'@Description("Warping constant")
Public Function Cw() As Double
End Function

'@Description("Horizontal principle axis moment of inertia")
Public Function Ix() As Double
End Function

'@Description("Vertical principle axis moment of inertia")
Public Function Iy() As Double
End Function

'@Description("Moment of inertia about rotated axis at a given angle from the horizontal principle axis")
Public Function Iu() As Double
End Function

'@Description("Moment of inertia about rotated axis at a given angle from the vertical principle axis")
Public Function Iv() As Double
End Function

'@Description("Polar moment of inertia")
Public Function Iz() As Double
End Function

'@Description("St. Venant Torsional constant")
Public Function J() As Double
End Function

'@Description("Horizontal axis radius of gyration")
Public Function Rx() As Double
End Function

'@Description("Vertical axis radius of gyration")
Public Function Ry() As Double
End Function

'@Description("Section modulus about horizontal axis")
Public Function Sx() As Double
End Function

'@Description("Section modulus about vertical axis")
Public Function Sy() As Double
End Function

'@Description("Plastic section modulus about horizontal axis")
Public Function Zx() As Double
End Function

'@Description("Plastic section modulus about vertical axis")
Public Function Zy() As Double
End Function
