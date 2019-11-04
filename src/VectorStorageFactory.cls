VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("StiffnessMethod.Model.Linear Algebra.Vector.Factory")

Option Explicit

Private Const CLASS_NAME As String = "VectorStorageFactory"

'@Description("Create a matrix storage factory by selecting the desired storage scheme")
Public Function CreateFactory(ByVal VectorType As VectorStorageScheme) As IVectorStorage

    Select Case VectorType
    
        Case VectorStorageScheme.Dense
            Set CreateFactory = New DenseVectorStorageFactory
            
        Case Else
            VectorErrors.ThrowStorageSchemeError CLASS_NAME
            
    End Select
    
End Function

'@Description("Create a matrix storage factory based on an existing matrix storage object")
Public Function CreateFactoryMatchingObject(ByVal VectorStorage As IVectorStorage) As IVectorStorageFactory

    If TypeOf VectorStorage Is DenseVectorStorage Then
        
        Set CreateFactoryMatchingObject = New DenseVectorStorageFactory
    
    Else
        
        MatrixErrors.ThrowStorageSchemeError CLASS_NAME
    
    End If

End Function
