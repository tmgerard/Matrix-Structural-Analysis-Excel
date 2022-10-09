Attribute VB_Name = "TestMatrix"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Equations")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
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
Private Sub TestIsSquare()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2)

    'Act:

    'Assert:
    Assert.IsTrue mat.IsSquare

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestIsNotSquare()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 3)

    'Act:

    'Assert:
    Assert.IsFalse mat.IsSquare

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestValueIsZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2)

    'Act:

    'Assert:
    Assert.AreEqual 0#, mat.ValueAt(0, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestLetGetValueAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2)

    'Act:
    Dim expected As Double
    expected = 10#
    
    mat.ValueAt(0, 1) = expected

    'Assert:
    Assert.AreEqual expected, mat.ValueAt(0, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestSetData()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matData(3) As Double
    matData(0) = 1
    matData(1) = 2
    matData(2) = 3
    matData(3) = 4
    
    Dim expected(3) As Double
    expected(0) = 1
    expected(1) = 12
    expected(2) = 3
    expected(3) = 4
    
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2).SetData(matData).AddTo(0, 1, 10)

    'Act:
    

    'Assert:
    AssertMatrixHasData mat, expected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestIdentityRow()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matData(3) As Double
    matData(0) = 2
    matData(1) = 3
    matData(2) = 4
    matData(3) = 5
    
    Dim expected(3) As Double
    expected(0) = 1
    expected(1) = 0
    expected(2) = 4
    expected(3) = 5
    
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2).SetData(matData).SetIdentityRow(0)

    'Act:
    

    'Assert:
    AssertMatrixHasData mat, expected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestIdentityColumn()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matData(3) As Double
    matData(0) = 2
    matData(1) = 3
    matData(2) = 4
    matData(3) = 5
    
    Dim expected(3) As Double
    expected(0) = 2
    expected(1) = 0
    expected(2) = 4
    expected(3) = 1
    
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2).SetData(matData).SetIdentityColumn(1)

    'Act:
    

    'Assert:
    AssertMatrixHasData mat, expected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Basic Operation")
Private Sub TestScale()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matData(3) As Double
    matData(0) = 1
    matData(1) = 2
    matData(2) = 3
    matData(3) = 4
    
    Dim expected(3) As Double
    expected(0) = 2
    expected(1) = 4
    expected(2) = 6
    expected(3) = 8
    
    Dim mat As Matrix
    Set mat = New Matrix
    Set mat = mat.SetSize(2, 2).SetData(matData).ScaleValues(2)

    'Act:
    

    'Assert:
    AssertMatrixHasData mat, expected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' Custom assert function to check values in matrix against values in array
Private Sub AssertMatrixHasData(ByRef mat As Matrix, ByRef matData() As Double)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim offset As Long
    For rowIndex = 0 To mat.Rows - 1
        offset = mat.Columns * rowIndex
        For colIndex = 0 To mat.Columns - 1
            Assert.AreEqual matData(offset + colIndex), mat.ValueAt(rowIndex, colIndex)
        Next colIndex
    Next rowIndex
    
End Sub

