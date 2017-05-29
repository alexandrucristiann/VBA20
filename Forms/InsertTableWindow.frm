VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertTableWindow 
   Caption         =   "Insert Table"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15675
   OleObjectBlob   =   "InsertTableWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertTableWindow"
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

Private Sub Insert_Click()
    ' check if table name is valid
    If TableName.Value = "" Or IsNumeric(TableName.Value) Then
        errorOut ("Cannot insert values into empty/invalid table name")
        Exit Sub
    End If
    
    ' check if columns are valid
    If Values.Value = "" Then
        errorOut ("Cannot insert empty values into table")
    End If
    
    Dim err As Boolean
    ' inert values into given table
    err = InsertTable(TableName.Value, Values.Value)
    If err = False Then
        errorOut ("error occured inserting values")
        Exit Sub
    End If
    
    ' unlod and exit everything if all went ok
    Me.Hide
    Unload Me
End Sub

Private Sub TableName_Change()
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
     If TableName.Value = "" Or IsNumeric(TableName.Value) Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
        Exit Sub
    End If
    
    Dim i As Integer
    Dim found As Boolean
    
    ' check if the tables specified exists
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).name = TableName.Value Then
            found = True
            Exit For
        End If
    Next i
    
    If found = False Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
        Exit Sub
    End If
End Sub

Private Sub Values_Change()
    ValuesLabel.ForeColor = vbBlack
    Values.BackColor = vbWhite
    
    Dim i As Integer
    Dim found As Boolean
    Dim sheet As Worksheet
    ' check if the tables specified exists
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).name = TableName.Value Then
            Set sheet = ActiveWorkbook.Worksheets(i)
            found = True
            Exit For
        End If
    Next i
    
    If found = False Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
        Exit Sub
    End If
End Sub
