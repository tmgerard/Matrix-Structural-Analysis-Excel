VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("StructuralAnalysis.Model.LinearAlgebra.Matrix.Factory")
Option Explicit

Private Const CLASS_NAME As String = "MatrixStorageFactory"

'@Description("Create a matrix storage factory by selecting the desired storage scheme")
Public Function CreateFactory(ByVal MatrixType As MatrixStorageScheme) As IMatrixStorageFactory

    Select Case MatrixType
    
        Case MatrixStorageScheme.DenseColumnMajor
            Set CreateFactory = New DenseColMajMatrixStorageFactory
            
        Case MatrixStorageScheme.DenseRowMajor
            Set CreateFactory = New DenseRowMajMatrixStorageFactory
            
        Case MatrixStorageScheme.DenseColumnVector
            Set CreateFactory = New DenseColumnVectorStorageFactory
            
        Case MatrixStorageScheme.DenseRowVector
            Set CreateFactory = New DenseRowVectorStorageFactory
            
        Case Else
            MatrixErrors.ThrowStorageSchemeError CLASS_NAME
            
    End Select
    
End Function

'@Description("Create a matrix storage factory based on an existing matrix storage object")
Public Function CreateFactoryMatchingObject(ByVal Storage As IMatrixStorage) As IMatrixStorageFactory

    If TypeOf Storage Is DenseColumnMajorMatrixStorage Then
        Set CreateFactoryMatchingObject = New DenseColMajMatrixStorageFactory
        
    ElseIf TypeOf Storage Is DenseRowMajorMatrixStorage Then
        Set CreateFactoryMatchingObject = New DenseRowMajMatrixStorageFactory
    
    ElseIf TypeOf Storage Is DenseColumnVectorStorage Then
        Set CreateFactoryMatchingObject = New DenseColumnVectorStorageFactory
        
    ElseIf TypeOf Storage Is DenseRowVectorStorage Then
        Set CreateFactoryMatchingObject = New DenseRowVectorStorageFactory
    
    Else
        MatrixErrors.ThrowStorageSchemeError CLASS_NAME
    
    End If

End Function

