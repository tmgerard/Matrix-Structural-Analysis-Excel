Attribute VB_Name = "DenseRowVectorFactoryTests"
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

Private factory As IMatrixStorageFactory
Private Const CREATE_LENGTH As Long = 10

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
    Set factory = New DenseRowVectorStorageFactory
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set factory = Nothing
End Sub

'@TestMethod("Factory")
Private Sub TestCreate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim vectorData As IVectorStorage

    'Act:
    Set vectorData = factory.Create(1, CREATE_LENGTH)

    'Assert:
    Assert.IsTrue TypeOf vectorData Is DenseRowVectorStorage

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

