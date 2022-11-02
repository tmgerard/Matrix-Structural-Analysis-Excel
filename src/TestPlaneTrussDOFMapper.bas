Attribute VB_Name = "TestPlaneTrussDOFMapper"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Model.Structure Model")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private node1 As Node2D
Private node2 As Node2D
Private node3 As Node2D
Private node4 As Node2D

Private nodeCollection As Collection

Dim mapper As PlaneTrussDOFMapper

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set nodeCollection = New Collection
    
    Set node1 = New Node2D
    With node1
        Set .Position = MakePoint2D(0, 0)
        .xConstrained = True
        .yConstrained = True
    End With
    
    nodeCollection.Add node1
    
    Set node2 = New Node2D
    With node2
        Set .Position = MakePoint2D(10, 0)
        .xConstrained = False
        .yConstrained = False
    End With
    
    nodeCollection.Add node2
    
    Set node3 = New Node2D
    With node3
        Set .Position = MakePoint2D(20, 0)
        .xConstrained = True
        .yConstrained = True
    End With
    
    nodeCollection.Add node3
    
    Set node4 = New Node2D
    With node4
        Set .Position = MakePoint2D(10, 10)
        .xConstrained = False
        .yConstrained = False
    End With
    
    nodeCollection.Add node4
    
    Set mapper = New PlaneTrussDOFMapper
    mapper.MapDofs nodeCollection
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set nodeCollection = Nothing
    Set node1 = Nothing
    Set node2 = Nothing
    Set node3 = Nothing
    Set node4 = Nothing
    
    Set mapper = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Operation")
Private Sub TestDOFMapper()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Matrix
    Set expected = New Matrix
    expected.SetSize 2, 4
    expected.ValueAt(0, 1) = 1
    expected.ValueAt(1, 1) = 2
    expected.ValueAt(0, 3) = 3
    expected.ValueAt(1, 3) = 4

    'Act:
    Dim actual As Matrix
    Set actual = mapper.DofMap

    'Assert:
    Assert.IsTrue (expected.Equals(actual))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
