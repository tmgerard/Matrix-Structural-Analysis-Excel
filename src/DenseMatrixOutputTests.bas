Attribute VB_Name = "DenseMatrixOutputTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private factory As IMatrixStorageFactory
Const CREATE_ROWS As Long = 3
Const CREATE_COLUMNS As Long = 3

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
    Set factory = New DenseRowMajMatrixStorageFactory
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
    Set factory = Nothing
End Sub

'@TestMethod("Output")
Public Sub TestToImmediateWindow()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixData As IMatrixStorage
    Set MatrixData = factory.CreateIdentity(CREATE_ROWS)
    
    Dim MatrixOutput As IMatrixOutput
    Set MatrixOutput = New DenseMatrixOutput

    'Act:
    MatrixOutput.ToImmediateWindow MatrixData

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output")
Public Sub TestToWorksheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MatrixData As IMatrixStorage
    Set MatrixData = factory.CreateIdentity(CREATE_ROWS)
    
    Dim MatrixOutput As IMatrixOutput
    Set MatrixOutput = New DenseMatrixOutput
    
    Dim WorksheetCell As Range
    Set WorksheetCell = MatrixTestSheet.Range("A10")

    'Act:
    MatrixOutput.ToWorksheet MatrixData, WorksheetCell

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output")
Public Sub TestToFile()
    On Error GoTo TestFail
    
    Const TOFILE_ROWS As Long = 25
    
    'Arrange:
    Dim MatrixData As IMatrixStorage
    Set MatrixData = factory.CreateIdentity(TOFILE_ROWS)
    
    Dim MatrixOutput As IMatrixOutput
    Set MatrixOutput = New DenseMatrixOutput

    'Act:
    MatrixOutput.ToFile MatrixData

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


