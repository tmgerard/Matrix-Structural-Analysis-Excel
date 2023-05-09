Attribute VB_Name = "InputValidation"
'@Folder("StructuralAnalysis.Model.Input")
Option Explicit

Public Enum InputErrors
    DuplicateNode = 600 + vbObjectError
    DuplicateMember
End Enum

Public Sub CheckForDuplicateNode(ByRef NodeTable As ListObject, ByRef ChangedRow As Long)

    Dim modifiedInput As Node2D
    With NodeTable.DataBodyRange
        ' ignoring constraints because we are only interested if the points are in the same location
        If IsEmpty(.Item(ChangedRow, 2)) Or IsEmpty(.Item(ChangedRow, 3)) Then
            Exit Sub
        End If
        
        Set modifiedInput = MakeNode2D(.Item(ChangedRow, 1), MakePoint2D(.Item(ChangedRow, 2), .Item(ChangedRow, 3)))
    
        Dim rowCount As Long
        rowCount = NodeTable.DataBodyRange.Rows.Count
        
        Dim row As Long
        Dim currentNode As Node2D
        For row = 1 To rowCount
            Set currentNode = MakeNode2D(.Item(row, 1), MakePoint2D(.Item(row, 2), .Item(row, 3)))

            If currentNode.Equals(modifiedInput) And Not currentNode.ID = modifiedInput.ID Then
                Err.Raise Number:=InputErrors.DuplicateNode, _
                          Description:="Nodes " & currentNode.ID & " and " & modifiedInput.ID & " are duplicates"
            End If
        Next row
    End With

End Sub

Public Sub CheckForDuplicateMember(ByRef MemberTable As ListObject, ByRef ChangedRow As Long)
    
    Dim modifiedInputMemberID As Long
    Dim modifiedInputStartNode As Long
    Dim modifiedInputEndNode As Long
    With MemberTable.DataBodyRange
        
        modifiedInputMemberID = .Item(ChangedRow, 1)
        modifiedInputStartNode = .Item(ChangedRow, 2)
        modifiedInputEndNode = .Item(ChangedRow, 3)

        If IsEmpty(modifiedInputStartNode) Or IsEmpty(modifiedInputEndNode) Then
            Exit Sub
        End If
    
        Dim rowCount As Long
        rowCount = MemberTable.DataBodyRange.Rows.Count
        
        Dim row As Long
        Dim currentMemberID As Long
        Dim currentStartNodeID As Long
        Dim currentEndNodeID As Long
        
        For row = 1 To rowCount
        
            currentMemberID = .Item(row, 1)
            currentStartNodeID = .Item(row, 2)
            currentEndNodeID = .Item(row, 3)
            
            ' check if start = end
            If currentStartNodeID = modifiedInputStartNode And _
                currentEndNodeID = modifiedInputEndNode And _
                Not currentMemberID = modifiedInputMemberID Then
                
                Err.Raise Number:=InputErrors.DuplicateMember, _
                          Description:="Members " & currentMemberID & " and " & modifiedInputMemberID & " are duplicates"
            End If
            
            ' check if flipped start and end nodes are equal
            If currentStartNodeID = modifiedInputEndNode And _
                currentEndNodeID = modifiedInputStartNode And _
                Not currentMemberID = modifiedInputMemberID Then
                
                Err.Raise Number:=InputErrors.DuplicateMember, _
                          Description:="Members " & currentMemberID & " and " & modifiedInputMemberID & " are duplicates"
            End If
            
        Next row
    End With
End Sub

