VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteTableWindow 
   Caption         =   "Delete Table"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   OleObjectBlob   =   "DeleteTableWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteTableWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Back_Click()
    ' Hide the current window
    ' and unload it from memory
    Me.Hide
    Unload Me
End Sub





Private Sub Delete_Click()
    If Me.ComboBoxName.Value = "" Then
        errorOut ("No table present in dba")
        Exit Sub
    End If
    
    Dim err As Boolean
    err = DeleteTable(Me.ComboBoxName.Value)
    If err = False Then
        errorOut ("Cannot delete table, table does not exist")
        Exit Sub
    End If
    
    Me.Hide
    Unload Me
End Sub


Private Sub TableName_Change()
    ' On every change in the TableName field
    ' check if we are dealing with valid charachters or not
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    
    If TableName.Value = "" Or _
    IsNumeric(TableName.Value) Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
    End If
End Sub



Private Sub Userform_Initialize()
    For i = Me.ComboBoxName.ListCount - 1 To 0 Step -1
        Me.ComboBoxName.RemoveItem i
    Next i
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).name <> "dba_start" Then
            Me.ComboBoxName.AddItem ActiveWorkbook.Worksheets(i).name
        End If
    Next i
End Sub
