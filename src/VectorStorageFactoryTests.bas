Attribute VB_Name = "VectorStorageFactoryTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Linear Algebra.Vector")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
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
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryDenseVectorStorage()
    On Error GoTo TestFail
    
    'Arrange:
    Dim factory As IVectorStorageFactory
    
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
Private Sub TestCreateFactoryMatchingDenseVectorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Storage As IVectorStorage
    Set Storage = New DenseColumnVectorStorage
    Storage.Length = 4
    
    Dim factory As IVectorStorageFactory
    
    'Act:
    Set factory = MatrixStorageFactory.CreateFactoryMatchingObject(Storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseColumnVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
