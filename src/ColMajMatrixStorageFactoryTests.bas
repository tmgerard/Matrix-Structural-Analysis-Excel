Attribute VB_Name = "ColMajMatrixStorageFactoryTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests.Linear Algebra.Matrix")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private factory As IMatrixStorageFactory
Private Const CREATE_ROWS As Long = 5
Private Const CREATE_COLUMNS As Long = 6

'@ModuleInitialize
Public Sub ModuleInitialize()
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
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
    Set factory = New DenseColMajMatrixStorageFactory
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
    Set factory = Nothing
End Sub

'@TestMethod("Factory")
Public Sub TestCreate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrixData As IMatrixStorage

    'Act:
    Set matrixData = factory.Create(CREATE_ROWS, CREATE_COLUMNS)

    'Assert:
    Assert.IsTrue TypeOf matrixData Is DenseColumnMajorMatrixStorage

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Public Sub TestCreateFromRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sheetRange As Range
    Set sheetRange = MatrixTestSheet.Range("A2:B4")
    
    Dim Cell As Range
    For Each Cell In sheetRange
        Cell.Value2 = Cell.row
    Next Cell
    
    Dim matrixData As IMatrixStorage

    'Act:
    Set matrixData = factory.CreateFromRange(sheetRange)

    'Assert:
    Dim row As Long
    Dim column As Long
    For row = 0 To matrixData.Rows - 1
        For column = 0 To matrixData.Columns - 1
            ' NOTE: Worksheet ranges are base 1 indexed
            Assert.AreEqual sheetRange.Cells(row + 1, column + 1).Value2, matrixData.Element(row, column)
        
        Next column
    Next row

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Public Sub TestCreateIdentity()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrixData As IMatrixStorage

    'Act:
    Set matrixData = factory.CreateIdentity(CREATE_ROWS)

    'Assert:
    Assert.AreEqual CREATE_ROWS, matrixData.Rows
    Assert.AreEqual CREATE_ROWS, matrixData.Columns
    
    Dim row As Long
    Dim column As Long
    For row = 0 To matrixData.Rows - 1
        For column = 0 To matrixData.Columns - 1
        
            If row = column Then
                Assert.AreEqual 1#, matrixData.Element(row, column)
            Else
                Assert.AreEqual 0#, matrixData.Element(row, column)
            End If
        
        Next column
    Next row

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Factory")
Public Sub TestCreateSquare()
    On Error GoTo TestFail
    
    'Arrange:
    Dim matrixData As IMatrixStorage

    'Act:
    Set matrixData = factory.CreateSquare(CREATE_ROWS)

    'Assert:
    Assert.AreEqual CREATE_ROWS, matrixData.Rows
    Assert.AreEqual CREATE_ROWS, matrixData.Columns

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



