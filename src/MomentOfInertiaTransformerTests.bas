Attribute VB_Name = "MomentOfInertiaTransformerTests"
Option Explicit
Option Private Module

Private Const Ix As Double = 16
Private Const Iy As Double = 9
Private Const Ixy As Double = 10
Private Const rotation As Double = 45

Private transform As MomentOfInertiaTransformer

'@TestModule
'@Folder("Tests.StructuralCrossSections.SectionProperties")

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
    Set transform = New MomentOfInertiaTransformer
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set transform = Nothing
End Sub

'@TestMethod("Calculation")
Private Sub TestIu()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 2.5

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, _
        actual:=transform.Iu(Ix, Iy, Ixy, WorksheetFunction.radians(rotation))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestIv()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 22.5

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, _
        actual:=transform.Iv(Ix, Iy, Ixy, WorksheetFunction.radians(rotation))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestIuv()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 3.5

    'Act:
    Dim actual As Double
    actual = transform.Iuv(Ix, Iy, Ixy, WorksheetFunction.radians(rotation))

    'Assert:
    Assert.AreEqual Expected:=Expected, _
        actual:=Round(actual, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

