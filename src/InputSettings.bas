Attribute VB_Name = "InputSettings"
'@Folder("StructuralAnalysis.Model.Input")
Option Explicit

' Table names in input worksheet
Public Const NodesTableName As String = "Nodes"
Public Const MembersTableName As String = "Members"
Public Const CrossSectionTableName As String = "CrossSections"
Public Const MaterialsTableName As String = "Materials"
Public Const LoadsTableName As String = "Loads"

' Constraint Conditions
Public Const Free As String = "FREE"
Public Const Pin As String = "PIN"
Public Const Roller As String = "ROLLER"
