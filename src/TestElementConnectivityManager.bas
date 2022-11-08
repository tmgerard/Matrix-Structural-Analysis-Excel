Attribute VB_Name = "TestElementConnectivityManager"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Model.Structure Model")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Dim dofMapStub As DOFMapperStub
Dim manager As ElementConnectivityManager
Dim elementCollection As Collection

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set dofMapStub = New DOFMapperStub
    
    Set elementCollection = New Collection
    elementCollection.Add MakeBarElement2D(1, _
        MakeNode2D(1, MakePoint2D(0, 0), True, True), _
        MakeNode2D(2, MakePoint2D(10, 0), False, False), _
        100, _
        100)
    
    Set manager = New ElementConnectivityManager
    Set manager.DofMap = dofMapStub
    Set manager.Elements = elementCollection
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set dofMapStub = Nothing
    Set elementCollection = Nothing
    Set manager = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Basic Operation")
Private Sub TestGetConnectivityArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(0 To 3) As Long
    expected(0) = 0
    expected(1) = 0
    expected(2) = 1
    expected(3) = 2

    'Act:
    Dim actual() As Long
    actual = manager.GetConnectivityArray(1)

    'Assert:
    Dim arrIndex As Long
    For arrIndex = LBound(expected) To UBound(expected)
        Assert.AreEqual expected(arrIndex), actual(arrIndex)
    Next arrIndex

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
