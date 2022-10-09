Attribute VB_Name = "TableNullValueChecker"
'@Folder("StructuralAnalysis.Model.Input")
Option Explicit

Public Enum InputTableErrors
    nullValue = 700 + vbObjectError
End Enum

Public Sub CheckTableForNull(ByRef table As ListObject)

    Dim rowCount As Long
    rowCount = table.DataBodyRange.Rows.count
    
    Dim columnCount As Long
    columnCount = table.DataBodyRange.Columns.count
    
    Dim row As Long
    Dim column As Long
    For row = 1 To rowCount
        For column = 1 To columnCount
            If IsEmpty(table.DataBodyRange.Item(row, column)) Then
                Err.Raise InputTableErrors.nullValue, "TableNullValueChecker.CheckTableForNull", _
                    "Input table """ & table.Name & """ contains empty cells."
            End If
        Next column
    Next row

End Sub
