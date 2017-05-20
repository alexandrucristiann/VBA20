VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateTableWindow 
   Caption         =   "Create Table"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14070
   OleObjectBlob   =   "CreateTableWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateTableWindow"
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


Private Sub ColumnCount_Change()
    ColumnCount.BackColor = vbWhite
    ColumnCountLabel.ForeColor = vbBlack

    ' we should also not change the state of the spin
    ' if the value is not numeric or is not in this interval
    ' [CoulmnCountSpin.Max, CoulmnCountSpin.Min]
    If Not IsNumeric(ColumnCount.Value) Or _
        (ColumnCount.Value > ColumnCountSpin.Max) Or _
        (ColumnCount.Value < ColumnCountSpin.Min) Then
        ColumnCount.BackColor = vbRed
        ColumnCountLabel.ForeColor = vbRed
        Exit Sub
    End If
    
    
    ' if the user wants to change the number of columns
    ' manually we need to update also the ColumnCountSpin
    ' value as well
    ColumnCountSpin.Value = ColumnCount.Value
End Sub


Private Sub Columns_Change()
    Columns.BackColor = vbWhite
    ColumnsLabel.ForeColor = vbBlack
    Dim limit As Integer
    limit = CInt(ColumnCount.Value)
    If Not validateColumns(Columns.Value, limit) Then
        Columns.BackColor = vbRed
        ColumnsLabel.ForeColor = vbRed
    End If
End Sub

Private Sub ColumnCountSpin_Change()
    ' when we increase or decrease the value
    ' update it in the field form
    ColumnCount.Value = ColumnCountSpin.Value
End Sub


' Create a new table in our database
' In our case a table is just a new sheet
Private Sub Create_Click()
    ' define here the default state
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    ColumnCountLabel.ForeColor = vbBlack
    ColumnCount.BackColor = vbWhite
    Columns.BackColor = vbWhite
    ColumnsLabel.ForeColor = vbBlack
    
    ' check all fields from the frame before everything else(bae)
    '
    ' table name check
    If TableName.Value = "" Or _
    IsNumeric(TableName.Value) Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
        errorOut ("Invalid table name")
        Exit Sub
    End If
    
    'column count check
     If Not IsNumeric(ColumnCount.Value) Or _
        (ColumnCount.Value > ColumnCountSpin.Max) Or _
        (ColumnCount.Value < ColumnCountSpin.Min) Then
        ColumnCount.BackColor = vbRed
        ColumnCountLabel.ForeColor = vbRed
        errorOut ("Invalid column count, count must be in [1,200]")
        Exit Sub
    End If
    
    'columns check
    Dim limit As Integer
    limit = CInt(ColumnCount.Value)
    If Not validateColumns(Columns.Value, limit) Then
        Columns.BackColor = vbRed
        ColumnsLabel.ForeColor = vbRed
        errorOut ("Invalid column names,length or found duplicates")
        Exit Sub
    End If
    
    'check if the table with the name passed already exists
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If TableName.Value = ActiveWorkbook.Worksheets(i).name Then
            TableNameLabel.ForeColor = vbRed
            TableName.BackColor = vbRed
            errorOut ("Table is already created")
            Exit Sub
        End If
    Next i
    
    ' create the table with the given columns
    Dim arr() As String
    arr = Split(Columns.Value, ",", limit)
    CreateTable TableName.Value, limit, arr
    
End Sub

Private Sub CreateTableFrame_Click()

End Sub

' On every change in the TableName field
' check if we are dealing with valid charachters or not
Private Sub TableName_Change()
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    
    If TableName.Value = "" Or _
    IsNumeric(TableName.Value) Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
    End If
End Sub

Private Sub UserForm_Initialize()
    ' when we open the create table form this initial state
    ' render the current value from the ColumnCountSpin
    ' into ColumnCount form field
    ColumnCount.Value = ColumnCountSpin.Value
End Sub
