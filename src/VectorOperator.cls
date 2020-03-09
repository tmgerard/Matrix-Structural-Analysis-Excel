VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralAnalysis.Math.LinearAlgebra.Vector.Operations")
Option Explicit

Private Const CLASS_NAME As String = "VectorOperator"

Public Function Add(ByVal leftVector As IVector, ByVal rightVector As IVector) As IVector

    If Not IsSameSize(leftVector, rightVector) Or _
       Not IsSameOrientation(leftVector, rightVector) Then
       
        VectorErrors.ThrowAdditionError CLASS_NAME
        
    End If
    
    Dim result As IVector
    Set result = cloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Storage.Element(index) = leftVector.Storage.Element(index) + rightVector.Storage.Element(index)
    Next index
    
    Set Add = result
    
End Function

Public Function CrossProduct(ByVal leftVector As IVector, ByVal rightVector As IVector) As IVector

    Const VECTOR_DIMENSION_CONSTRAINT As Long = 3

    If Not rightVector.Length = VECTOR_DIMENSION_CONSTRAINT Or _
       Not leftVector.Length = VECTOR_DIMENSION_CONSTRAINT Then
       
        VectorErrors.ThrowCrossProductError CLASS_NAME
        
    End If
    
    Dim result As IVector
    Set result = cloneVector(leftVector)
    
    With result.Storage
        .Element(0) = (leftVector.Storage.Element(1) * rightVector.Storage.Element(2)) _
            - (leftVector.Storage.Element(2) * rightVector.Storage.Element(1))
        .Element(1) = (leftVector.Storage.Element(2) * rightVector.Storage.Element(0)) _
            - (leftVector.Storage.Element(0) * rightVector.Storage.Element(2))
        .Element(2) = (leftVector.Storage.Element(0) * rightVector.Storage.Element(1)) _
            - (leftVector.Storage.Element(1) * rightVector.Storage.Element(0))
    End With
    
    Set CrossProduct = result
    
End Function

Public Function DotProduct(ByVal leftVector As IVector, ByVal rightVector As IVector) As Double

    If Not IsSameSize(leftVector, rightVector) Then
       
        VectorErrors.ThrowDotProductError CLASS_NAME
        
    End If
    
    Dim result As Double
    result = 0
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result = result + leftVector.Storage.Element(index) * rightVector.Storage.Element(index)
    Next index
    
    DotProduct = result
    
End Function

Public Function EuclideanDistance(ByVal vector As IVector) As Double
    
    Dim result As Double
    result = 0
    
    Dim index As Long
    For index = 0 To vector.Length - 1
        result = result + vector.Storage.Element(index) * vector.Storage.Element(index)
    Next index
    
    EuclideanDistance = Math.Sqr(result)
    
End Function

Public Function Equals(ByVal leftVector As IVector, ByVal rightVector As IVector) As Boolean

    Dim result As Boolean
    result = True
    
    If Not leftVector Is rightVector Then
        result = False
        Exit Function
    End If
    
    If Not leftVector.Length = rightVector.Length Then
        result = False
        Exit Function
    End If
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        If Not leftVector.Storage.Element(index) = rightVector.Storage.Element(index) Then
            result = False
            Exit For
        End If
    Next index
    
    Equals = result

End Function

Public Function ScalarMultiply(ByVal leftVector As IVector, ByVal scalarValue As Double) As IVector
    
    Dim result As IVector
    Set result = cloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Storage.Element(index) = leftVector.Storage.Element(index) * scalarValue
    Next index
    
    Set ScalarMultiply = result
    
End Function

Public Function Subtract(ByVal leftVector As IVector, ByVal rightVector As IVector) As IVector

    If Not IsSameSize(leftVector, rightVector) Or _
       Not IsSameOrientation(leftVector, rightVector) Then
       
        VectorErrors.ThrowSubtractionError CLASS_NAME
        
    End If
    
    Dim result As IVector
    Set result = cloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Storage.Element(index) = leftVector.Storage.Element(index) - rightVector.Storage.Element(index)
    Next index
    
    Set Subtract = result
    
End Function

Private Function cloneVector(ByVal vectorToClone As IVector) As IVector
   
   Dim newVector As ICloneable
    Set newVector = vectorToClone
    
    Set cloneVector = newVector.Clone
    
End Function

Private Function IsSameSize(ByVal leftVector As IVector, ByVal rightVector As IVector) As Boolean
    
    IsSameSize = Not (Not leftVector.Length = rightVector.Length)
    
End Function

Private Function IsSameOrientation(ByVal leftVector As IVector, ByVal rightVector As IVector) As Boolean
    
    IsSameOrientation = Not (Not leftVector.Orientation = rightVector.Orientation)
    
End Function


