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
    Set factory = VectorStorageFactory.CreateFactory(VectorStorageScheme.Dense)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFactoryMatchingDenseVectorStorageObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim storage As IVectorStorage
    Set storage = New DenseVectorStorage
    storage.Length = 4
    
    Dim factory As IVectorStorageFactory
    
    'Act:
    Set factory = VectorStorageFactory.CreateFactoryMatchingObject(storage)

    'Assert:
    Assert.IsTrue TypeOf factory Is DenseVectorStorageFactory

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
