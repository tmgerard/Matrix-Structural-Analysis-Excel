VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Model.StructuralCrossSections.Shapes")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements IStructuralCrossSection

Private Type TRectangle
    Width As Double
    Height As Double
    rotation As Double
End Type

Private this As TRectangle
Private Const CLASS_NAME As String = "Rectangle"

Public Sub Create(ByVal Width As Double, ByVal Height As Double, Optional ByVal rotation As Double = 0)
    
    CheckDimension Width
    CheckDimension Height
    
    With this
        .Width = Width
        .Height = Height
        .rotation = rotation
    End With
    
End Sub

Public Property Get Height() As Double
    Height = this.Height
End Property

Public Property Let Height(ByVal value As Double)
    CheckDimension value
    this.Height = value
End Property

Public Property Get rotation() As Double
    rotation = this.rotation
End Property

Public Property Let rotation(ByVal radians As Double)
    this.rotation = radians
End Property

Public Property Get Width() As Double
    Width = this.Width
End Property

Public Property Let Width(ByVal value As Double)
    CheckDimension value
    this.Width = value
End Property

'@Description("Cross sectional area")
Public Function Area() As Double
    Area = this.Width * this.Height
End Function

'@Description("Warping constant")
' The warping constant for closed sections is often taken as zero
Public Function Cw() As Double
    Cw = 0
End Function

'@Description("Horizontal principle axis moment of inertia")
Public Function Ix() As Double
    Ix = this.Width * this.Height ^ 3 / 12#
End Function

'@Description("Vertical principle axis moment of inertia")
Public Function Iy() As Double
    Iy = this.Height * this.Width ^ 3 / 12#
End Function

'@Description("Product of inertia")
Public Function Ixy() As Double
    Ixy = 0
End Function

'@Description("Moment of inertia about rotated axis at a given angle from the horizontal principle axis")
Public Function Iu() As Double
    Dim transform As MomentOfInertiaTransformer
    Set transform = New MomentOfInertiaTransformer
    
    Iu = transform.Iu(Me.Ix, Me.Iy, Me.Ixy, this.rotation)
End Function

'@Description("Moment of inertia about rotated axis at a given angle from the vertical principle axis")
Public Function Iv() As Double
    Dim transform As MomentOfInertiaTransformer
    Set transform = New MomentOfInertiaTransformer
    
    Iv = transform.Iv(Me.Ix, Me.Iy, Me.Ixy, this.rotation)
End Function

'@Description("Product of inertia for rotated cross-section")
Public Function Iuv() As Double
    Dim transform As MomentOfInertiaTransformer
    Set transform = New MomentOfInertiaTransformer
    
    Iuv = transform.Iuv(Me.Ix, Me.Iy, Me.Ixy, this.rotation)
End Function

'@Description("Polar moment of inertia")
Public Function Iz() As Double
    Iz = Me.Ix + Me.Iy
End Function

'@Description("St. Venant torsional constant")
' Refer to Roark's Formulas for Stress and Strain 8th edition
' Table 10.7 for torsion constant formula of a rectangular cross-section
Public Function J() As Double
    Dim a As Double
    a = 0.5 * WorksheetFunction.Max(this.Height, this.Width)
    
    Dim b As Double
    b = 0.5 * WorksheetFunction.Min(this.Height, this.Width)
    
    Dim result As Double
    result = 3.36 * (b / a) * (1 - (b ^ 4 / (12# * a ^ 4)))
    result = 16# / 3# - result
    result = (a * b ^ 3) * result
    
    J = result
End Function

'@Description("Horizontal axis radius of gyration")
Public Function Rx() As Double
    Rx = Math.Sqr(Me.Ix / Me.Area)
End Function

'@Description("Vertical axis radius of gyration")
Public Function Ry() As Double
    Ry = Math.Sqr(Me.Iy / Me.Area)
End Function

'@Description("Section modulus about horizontal axis")
Public Function Sx() As Double
    Sx = this.Width * this.Height ^ 2 / 6#
End Function

'@Description("Section modulus about vertical axis")
Public Function Sy() As Double
    Sy = this.Height * this.Width ^ 2 / 6#
End Function

'@Description("Plastic section modulus about horizontal axis")
Public Function Zx() As Double
    Zx = this.Width * this.Width ^ 2 / 4#
End Function

'@Description("Plastic section modulus about vertical axis")
Public Function Zy() As Double
    Zy = this.Height * this.Width ^ 2 / 4#
End Function

Private Sub CheckDimension(ByVal dimension As Double)
    If dimension <= 0 Then StructuralCrossSectionErrors.ThrowDimensionError CLASS_NAME
End Sub

' IStructuralCrossSection Implementation
Private Function IStructuralCrossSection_Area() As Double
    IStructuralCrossSection_Area = Me.Area
End Function

Private Function IStructuralCrossSection_Cw() As Double
    IStructuralCrossSection_Cw = Me.Cw
End Function

Private Function IStructuralCrossSection_Iu() As Double
    IStructuralCrossSection_Iu = Me.Iu
End Function

Private Function IStructuralCrossSection_Iv() As Double
    IStructuralCrossSection_Iv = Me.Iv
End Function

Private Function IStructuralCrossSection_Ix() As Double
    IStructuralCrossSection_Ix = Me.Ix
End Function

Private Function IStructuralCrossSection_Iy() As Double
    IStructuralCrossSection_Iy = Me.Iy
End Function

Private Function IStructuralCrossSection_Iz() As Double
    IStructuralCrossSection_Iz = Me.Iz
End Function

Private Function IStructuralCrossSection_J() As Double
    IStructuralCrossSection_J = Me.J
End Function

Private Function IStructuralCrossSection_Rx() As Double
    IStructuralCrossSection_Rx = Me.Rx
End Function

Private Function IStructuralCrossSection_Ry() As Double
    IStructuralCrossSection_Ry = Me.Ry
End Function

Private Function IStructuralCrossSection_Sx() As Double
    IStructuralCrossSection_Sx = Me.Sx
End Function

Private Function IStructuralCrossSection_Sy() As Double
    IStructuralCrossSection_Sy = Me.Sy
End Function

Private Function IStructuralCrossSection_Zx() As Double
    IStructuralCrossSection_Zx = Me.Zx
End Function

Private Function IStructuralCrossSection_Zy() As Double
    IStructuralCrossSection_Zy = Me.Zy
End Function
