Attribute VB_Name = "MatrixStorageFactoryTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests.LinearAlgebra.Factory")
'@IgnoreModule

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        'Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New AssertClass
        'Set Fakes = New FakesProvider
    #End If
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryDenseColumnMajor()
    On Error GoTo TestFail
    
    'Arrange:
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactory(MatrixStorageScheme.DenseColumnMajor)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseColMajMatrixStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryDenseRowMajor()
    On Error GoTo TestFail
    
    'Arrange:
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactory(MatrixStorageScheme.DenseRowMajor)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseRowMajMatrixStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryDenseRowVector()
    On Error GoTo TestFail
    
    'Arrange:
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactory(MatrixStorageScheme.DenseRowVector)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseRowVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryDenseColumnVector()
    On Error GoTo TestFail
    
    'Arrange:
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactory(MatrixStorageScheme.DenseColumnVector)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseColumnVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryMatchingDenseColumnMajorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Storage As IMatrixStorage
    Set Storage = New DenseColumnMajorMatrixStorage
    Storage.SetSize Rows:=2, Columns:=2
    
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(Storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseColMajMatrixStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryMatchingDenseRowMajorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Storage As IMatrixStorage
    Set Storage = New DenseRowMajorMatrixStorage
    Storage.SetSize Rows:=2, Columns:=2
    
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(Storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseRowMajMatrixStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryMatchingDenseRowVectorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Storage As IMatrixStorage
    Set Storage = New DenseRowVectorStorage
    Storage.SetSize Rows:=1, Columns:=4
    
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(Storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseRowVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryMatchingDenseColumnVectorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Storage As IMatrixStorage
    Set Storage = New DenseColumnVectorStorage
    Storage.SetSize Rows:=4, Columns:=1
    
    Dim factory As IMatrixStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(Storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseColumnVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
