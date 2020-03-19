Attribute VB_Name = "PointLoadFEMTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.FixedEndForces")

Private Const pLoad As Double = 10#
Private Const lengthOfBeam As Double = 16#
Private Const aLength As Double = 6#
Private Const bLength As Double = 10#
Private femCalculator As PointLoadFEM


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
    Set femCalculator = New PointLoadFEM
    With femCalculator
        .BeamLength = lengthOfBeam
        .BeamLoad = pLoad
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set femCalculator = Nothing
End Sub

'@TestMethod("Calculation")
Private Sub TestLeftFixedEndMomentWithLoadAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 20

    'Act:
    femCalculator.DistanceToLoad = lengthOfBeam / 2

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndMoment(Left)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRightFixedEndMomentWithLoadAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 20

    'Act:
    femCalculator.DistanceToLoad = lengthOfBeam / 2

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndMoment(Right)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestLeftFixedEndReactionWithLoadAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 5

    'Act:
    femCalculator.DistanceToLoad = lengthOfBeam / 2

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndReaction(Left)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRightFixedEndReactionWithLoadAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 5

    'Act:
    femCalculator.DistanceToLoad = lengthOfBeam / 2

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndReaction(Right)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestLeftFixedEndMomentWithLoadNotAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 23.4375

    'Act:
    femCalculator.DistanceToLoad = aLength

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndMoment(Left)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRightFixedEndMomentWithLoadNotAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 14.0625

    'Act:
    femCalculator.DistanceToLoad = aLength

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndMoment(Right)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestLeftFixedEndReactionWithLoadNotAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 6.8359375

    'Act:
    femCalculator.DistanceToLoad = aLength

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndReaction(Left)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestRightFixedEndReactionWithLoadNotAtMidPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedFEM As Double = 3.1640625

    'Act:
    femCalculator.DistanceToLoad = aLength

    'Assert:
    Assert.AreEqual expectedFEM, femCalculator.FixedEndReaction(Right)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestZeroBeamLength()
    Const ExpectedError As Long = StructuralModelError.BadElementLength
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    femCalculator.BeamLength = 0

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
Private Sub TestNegativeBeamLength()
    Const ExpectedError As Long = StructuralModelError.BadElementLength
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    femCalculator.BeamLength = -1

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
Private Sub TestBadLoadLocationOffLeftSide()
    Const ExpectedError As Long = StructuralModelError.BadElementLoadLocation
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    femCalculator.DistanceToLoad = -1

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
Private Sub TestBadLoadLocationOffRightSide()
    Const ExpectedError As Long = StructuralModelError.BadElementLoadLocation
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    femCalculator.DistanceToLoad = 10000

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
