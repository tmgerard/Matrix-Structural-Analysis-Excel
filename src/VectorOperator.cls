VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StiffnessMethod.Model.Linear Algebra.Vector.Operations")
Option Explicit

Private Const CLASS_NAME As String = "VectorOperator"

Public Function Add(ByVal leftVector As IVector, ByVal rightVector As IVector) As IVector

    If Not IsSameSize(leftVector, rightVector) Or _
       Not IsSameOrientation(leftVector, rightVector) Then
       
        VectorErrors.ThrowAdditionError CLASS_NAME
        
    End If
    
    Dim result As IVector
    Set result = CloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Element(index) = leftVector.Element(index) + rightVector.Element(index)
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
    Set result = CloneVector(leftVector)
    
    With result
        .Element(0) = (leftVector.Element(1) * rightVector.Element(2)) - (leftVector.Element(2) * rightVector.Element(1))
        .Element(1) = (leftVector.Element(2) * rightVector.Element(0)) - (leftVector.Element(0) * rightVector.Element(2))
        .Element(2) = (leftVector.Element(0) * rightVector.Element(1)) - (leftVector.Element(1) * rightVector.Element(0))
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
        result = result + leftVector.Element(index) * rightVector.Element(index)
    Next index
    
    DotProduct = result
    
End Function

Public Function EuclideanDistance(ByVal vector As IVector) As Double
    
    Dim result As Double
    result = 0
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result = result + vector.Element(index) * vector.Element(index)
    Next index
    
    EuclideanDistance = Math.Sqr(result)
    
End Function

Public Function ScalarMultiply(ByVal leftVector As IVector, ByVal scalarValue As Double) As IVector
    
    Dim result As IVector
    Set result = CloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Element(index) = leftVector.Element(index) * scalarValue
    Next index
    
    Set Subtract = result
    
End Function

Public Function Subtract(ByVal leftVector As IVector, ByVal rightVector As IVector) As IVector

    If Not IsSameSize(leftVector, rightVector) Or _
       Not IsSameOrientation(leftVector, rightVector) Then
       
        VectorErrors.ThrowSubtractionError CLASS_NAME
        
    End If
    
    Dim result As IVector
    Set result = CloneVector(leftVector)
    
    Dim index As Long
    For index = 0 To leftVector.Length - 1
        result.Element(index) = leftVector.Element(index) - rightVector.Element(index)
    Next index
    
    Set Subtract = result
    
End Function

Private Function CloneVector(ByVal vectorToClone As IVector) As IVector
   
   Dim newVector As ICloneable
    Set newVector = vectorToClone
    
    Set CloneVector = newVector.Clone
    
End Function

Private Function IsSameSize(ByVal leftVector As IVector, ByVal rightVector As IVector) As Boolean
    
    If Not leftVector.Length = rightVector.Length Then
        IsSameSize = False
    Else
        IsSameSize = True
    End If
    
End Function

Private Function IsSameOrientation(ByVal leftVector As IVector, ByVal rightVector As IVector) As Boolean
    
    If Not leftVector.Orientation = rightVector.Orientation Then
        IsSameOrientation = False
    Else
        IsSameOrientation = True
    End If
    
End Function
