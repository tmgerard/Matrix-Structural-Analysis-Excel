VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixStorageFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Model.Matrix.Factory")
Option Explicit

Private Const CLASS_NAME As String = "MatrixStorageFactory"

'@Description("Create a matrix storage factory by selecting the desired storage scheme")
Public Function CreateFactory(ByVal MatrixType As MatrixStorageScheme) As IMatrixStorageFactory

    Select Case MatrixType
    
        Case MatrixStorageScheme.DenseColumnMajor
            Set CreateFactory = New DenseColMajMatrixStorageFactory
            
        Case MatrixStorageScheme.DenseRowMajor
            Set CreateFactory = New DenseRowMajMatrixStorageFactory
            
        Case Else
            MatrixErrors.ThrowStorageSchemeError CLASS_NAME
            
    End Select
    
End Function

'@Description("Create a matrix storage factory based on an existing matrix storage object")
Public Function CreateFactoryMatchingObject(ByVal MatrixStorage As IMatrixStorage) As IMatrixStorageFactory

    Select Case MatrixStorage
    
        Case TypeOf MatrixStorage Is DenseColumnMajorMatrixStorage
            Set CreateFactoryMatchingObject = New DenseColMajMatrixStorageFactory
        
        Case TypeOf MatrixStorage Is DenseRowMajorMatrixStorage
            Set CreateFactoryMatchingObject = New DenseRowMajMatrixStorageFactory
        
        Case Else
            MatrixErrors.ThrowStorageSchemeError CLASS_NAME
    
    End Select

End Function

