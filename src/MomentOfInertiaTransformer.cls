VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MomentOfInertiaTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.StructuralCrossSections.SectionProperties")
Option Explicit

'@Description("Moment of inertia about rotated axis at a given angle from the horizontal axis")
'@Ignore UseMeaningfulName
Public Function Iu(ByVal Ix As Double, ByVal Iy As Double, ByVal Ixy As Double, ByVal radians As Double)

    Dim a As Double
    a = (Ix + Iy) / 2
    
    Dim b As Double
    b = ((Ix - Iy) / 2) * Math.Cos(2 * radians)
    
    Dim c As Double
    c = Ixy * Math.Sin(2 * radians)
    
    Iu = a + b - c

End Function

'@Description("Moment of inertia about rotated axis at a given angle from the vertical axis")
'@Ignore UseMeaningfulName
Public Function Iv(ByVal Ix As Double, ByVal Iy As Double, ByVal Ixy As Double, ByVal radians As Double)

    Dim a As Double
    a = (Ix + Iy) / 2
    
    Dim b As Double
    b = ((Ix - Iy) / 2) * Math.Cos(2 * radians)
    
    Dim c As Double
    c = Ixy * Math.Sin(2 * radians)
    
    Iv = a - b + c

End Function

'@Description("Product of inertia about rotated axis")
'Ignore UseMeaningfulName
Public Function Iuv(ByVal Ix As Double, ByVal Iy As Double, ByVal Ixy As Double, ByVal radians As Double)
    
    Dim a As Double
    a = ((Ix - Iy) / 2) * Math.Sin(2 * radians)
    
    Dim b As Double
    b = Ixy * Math.Cos(2 * radians)
    
    Iuv = a + b
    
End Function
