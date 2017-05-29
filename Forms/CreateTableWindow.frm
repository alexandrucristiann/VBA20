VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateTableWindow 
   Caption         =   "Create Table"
   ClientHeight    =   8220
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
Private Sub Add_Click()
    Column.BackColor = vbWhite
    ColumnLabel.ForeColor = vbBlack
    If Column.Value = "" Or IsNumeric(Column.Value) Then
        Column.BackColor = vbRed
        ColumnsLabel.ForeColor = vbRed
        errorOut ("Invalid column name")
        Exit Sub
    End If
    
    'add column to listbox
    Me.Columns.AddItem (Column.Value)
    Me.Columns.list(Me.Columns.ListCount - 1, 1) = Me.ComboBoxType.Value
End Sub

Private Sub Back_Click()
    ' Hide the current window
    ' and unload it from memory
    Me.Hide
    Unload Me
End Sub

Private Sub Column_Change()
    Column.BackColor = vbWhite
    ColumnLabel.ForeColor = vbBlack
    If Column.Value = "" Or IsNumeric(Column.Value) Then
        Column.BackColor = vbRed
        ColumnLabel.ForeColor = vbRed
    End If
End Sub


' Create a new table in our database
' In our case a table is just a new sheet
Private Sub Create_Click()
    ' define here the default state
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    
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
    
    'check if the table with the name passed already exists
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If TableName.Value = ActiveWorkbook.Worksheets(i).name Then
            TableNameLabel.ForeColor = vbRed
            TableName.BackColor = vbRed
            errorOut ("Table is already created")
            Exit Sub
        End If
    Next i
    
    ' append all column names and their types
    Dim columnTypes() As String
    Dim columnNames() As String
    Dim ncolumns As Long
    ncolumns = 0
    ReDim Preserve columnTypes(ncolumns)
    ReDim Preserve columnNames(ncolumns)
    For i = 0 To Columns.ListCount - 1
        ReDim Preserve columnTypes(0 To ncolumns)
        ReDim Preserve columnNames(0 To ncolumns)
        If Columns.Column(0, i) = "" Then
            errorOut ("column does not have type")
            Exit Sub
        End If
        columnTypes(ncolumns) = Columns.Column(1, i)
        columnNames(ncolumns) = Columns.Column(0, i)
        ncolumns = ncolumns + 1
    Next i
    
    Dim dba_start As Worksheet
    Set dba_start = ActiveWorkbook.Worksheets("dba_start")
    
    Dim emptyRow As Long
    emptyRow = dba_start.Cells(dba_start.Rows.Count, "A").End(xlUp).Row + 1 ' get the last raw that is empty
    If dba_start.Cells(1, "A").Value = "" Then
        emptyRow = 1
    End If
                
    dba_start.Cells(emptyRow, 1).Value = TableName.Value
    For i = 1 To ncolumns
        dba_start.Cells(emptyRow, i + 1).Value = columnTypes(i - 1)
    Next i
    
    CreateTable TableName.Value, columnNames
    
    Me.Hide
    Unload Me
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

Private Sub Userform_Initialize()
    ' lock the combo box type
    Me.ComboBoxType.Style = fmStyleDropDownList
    For i = Me.ComboBoxType.ListCount - 1 To 0 Step -1
        Me.ComboBoxType.RemoveItem i
    Next i
    ' Add types to chose from when creating a new table
    Me.ComboBoxType.AddItem ("String")
    Me.ComboBoxType.AddItem ("Integer")
End Sub
