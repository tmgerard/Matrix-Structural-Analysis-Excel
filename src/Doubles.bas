Attribute VB_Name = "Doubles"
'@Folder("StructuralAnalysis.Utilities")
Option Explicit

Private Const CompareTolerance As Double = 0.00001

'@Description "Compares two doubles with a given tolerance for equality."
Public Function Equal(ByRef num1 As Double, ByRef num2 As Double, _
    Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute Equal.VB_Description = "Compares two doubles with a given tolerance for equality."
    
    Equal = Math.Abs(num2 - num1) < tolerance

End Function

'@Description "Checks if given double is within a given tolerance of zero."
Public Function IsEffectivelyZero(ByRef num As Double, Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute IsEffectivelyZero.VB_Description = "Checks if given double is within a given tolerance of zero."
    IsEffectivelyZero = Equal(num, 0#, tolerance)
End Function

'@Description "Checks if a given double is within a given tolerance of one."
Public Function IsEffectivelyOne(ByRef num As Double, Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute IsEffectivelyOne.VB_Description = "Checks if a given double is within a given tolerance of one."
    IsEffectivelyOne = Equal(num, 1#, tolerance)
End Function
