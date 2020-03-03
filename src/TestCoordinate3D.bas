Attribute VB_Name = "TestCoordinate3D"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Geometry.Coordinates")

' coordinate point 1
Private Const X1 As Double = 1
Private Const Y1 As Double = 1
Private Const Z1 As Double = 1
Private point1 As Coordinate3d

' coordinate point 2
Private Const X2 As Double = 2
Private Const Y2 As Double = 2
Private Const Z2 As Double = 2
Private point2 As Coordinate3d

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
    Set point1 = New Coordinate3d
    point1.SetCoordinates X1, Y1, Z1
    
    Set point2 = New Coordinate3d
    point2.SetCoordinates X2, Y2, Z2
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set point1 = Nothing
    Set point2 = Nothing
End Sub

'@TestMethod("Calculation")
Private Sub TestDistance()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 1.73

    'Act:

    'Assert:
    Assert.AreEqual expected, Round(point1.Distance(point2), 2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
