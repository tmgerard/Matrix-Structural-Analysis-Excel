Attribute VB_Name = "DenseRowVectorFactoryTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Linear Algebra.Factory")

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
    Dim VectorData As IVectorStorage

    'Act:
    Set VectorData = factory.Create(1, CREATE_LENGTH)

    'Assert:
    Assert.IsTrue TypeOf VectorData Is DenseRowVectorStorage

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFromRangeRowOriented()
    On Error GoTo TestFail
    
    'Arrange:
    Dim VectorData As DenseRowVectorStorage
    
    Dim sheetRange As Range
    Set sheetRange = VectorTestSheet.Range("A2:D2")
    
    Dim Cell As Range
    For Each Cell In sheetRange
        Cell.Value2 = Cell.column
    Next Cell
    
    'Act:
    Set VectorData = factory.CreateFromRange(sheetRange)

    'Assert:
    Dim index As Long
    For index = 0 To VectorData.Length - 1
        Assert.AreEqual CDbl(sheetRange.Cells(sheetRange.row, index + 1).Value2), VectorData.Element(index)
    Next index

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Private Sub TestCreateFromRangeColumnOriented()
    On Error GoTo TestFail
    
    'Arrange:
    Dim VectorData As DenseRowVectorStorage
    
    Dim sheetRange As Range
    Set sheetRange = VectorTestSheet.Range("A5:A8")
    
    Dim Cell As Range
    For Each Cell In sheetRange
        Cell.Value2 = Cell.row
    Next Cell
    
    'Act:
    Set VectorData = factory.CreateFromRange(sheetRange)

    'Assert:
    Dim index As Long
    For index = 0 To VectorData.Length - 1
        Assert.AreEqual CDbl(sheetRange.Cells(index + 1, sheetRange.column).Value2), VectorData.Element(index)
    Next index

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestCreateFromRangeBadRange()
    Const ExpectedError As Long = VectorError.CreateFromRange
    On Error GoTo TestFail
    
    'Arrange:
    Dim VectorData As DenseRowVectorStorage
    
    Dim sheetRange As Range
    Set sheetRange = VectorTestSheet.Range("A5:B8") ' column or row count must be 1

    'Act:
    Set VectorData = factory.CreateFromRange(sheetRange)

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

