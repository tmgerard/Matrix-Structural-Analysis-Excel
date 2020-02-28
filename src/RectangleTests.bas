Attribute VB_Name = "RectangleTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.StructuralCrossSections.Shapes")

Private Const rectWidth As Double = 4#
Private Const rectHeight As Double = 3#
Private Const noRotation As Double = 0#
Private Const rotation As Double = 45 ' degrees

Private plate As Rectangle

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
    Set plate = New Rectangle
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set plate = Nothing
End Sub

'@TestMethod("Expected Error")
Private Sub TestCreate_NegativeWidth()
    Const ExpectedError As Long = CrossSectionError.dimension
    On Error GoTo TestFail
    
    'Arrange:
    plate.Create Width:=-rectWidth, Height:=rectHeight

    'Act:

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Private Sub TestCreate_NegativeHeight()
    Const ExpectedError As Long = CrossSectionError.dimension
    On Error GoTo TestFail
    
    'Arrange:
    plate.Create Width:=rectWidth, Height:=-rectHeight

    'Act:

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property")
Private Sub TestHeightProperty()
    On Error GoTo TestFail
    
    'Arrange:
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=rectHeight, actual:=plate.Height

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestWidthProperty()
    On Error GoTo TestFail
    
    'Arrange:
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=rectWidth, actual:=plate.Width

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestArea()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 12
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Area

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestIx()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 9
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Ix

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestIy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 16
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestIz()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 25
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Iz

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestSx()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 6
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestSy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 8
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestZx()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 9
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Zx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestZy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 12
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=plate.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRx()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 0.866025
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=Round(plate.Rx, 6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 1.154701
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=Round(plate.Ry, 6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestJ()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Double
    Expected = 19.438506
    
    plate.Create Width:=rectWidth, Height:=rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected:=Expected, actual:=Round(plate.J, 6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

