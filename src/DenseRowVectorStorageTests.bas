Attribute VB_Name = "DenseRowVectorStorageTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Linear Algebra.Matrix Storage")

Private Const VECTOR_LENGTH As Long = 4
Private Storage As DenseRowVectorStorage

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
    Set Storage = New DenseRowVectorStorage
    Storage.Length = VECTOR_LENGTH
    
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
    Set Storage = Nothing
    
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

'@TestMethod("Property")
Private Sub TestGetLength()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual VECTOR_LENGTH, Storage.Length

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Operation")
Private Sub TestClear()
    On Error GoTo TestFail
    
    'Arrange:
    Const EXPECTED_VALUE As Double = 0#

    'Act:
    With Storage
        .Element(0) = 1
        .Element(1) = 2
        .Element(2) = 3
        .Element(3) = 4
    End With
    
    Storage.Clear

    'Assert:
    Assert.AreEqual EXPECTED_VALUE, Storage.Element(0)
    Assert.AreEqual EXPECTED_VALUE, Storage.Element(1)
    Assert.AreEqual EXPECTED_VALUE, Storage.Element(2)
    Assert.AreEqual EXPECTED_VALUE, Storage.Element(3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Operation")
Private Sub TestClone()
    On Error GoTo TestFail
    
    'Arrange:
    Dim newStorage As DenseRowVectorStorage

    'Act:
    Set newStorage = Storage.Clone

    'Assert:
    Assert.IsTrue Not ObjPtr(Storage) = ObjPtr(newStorage)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


