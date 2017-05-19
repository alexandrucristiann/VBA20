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
    ' [base.MIN_COLUMNS, base.MAX_COLUMNS]
    If Not IsNumeric(ColumnCount.Value) Or _
        (ColumnCount.Value > CoulmnCountSpin.Max) Or _
        (ColumnCount.Value < CoulmnCountSpin.Min) Then
        ColumnCount.BackColor = vbRed
        ColumnCountLabel.ForeColor = vbRed
        Exit Sub
    End If
    
    
    ' if the user wants to change the number of columns
    ' manually we need to update also the ColumnCountSpin
    ' value as well
    CoulmnCountSpin.Value = ColumnCount.Value
End Sub

Private Sub Columns_Change()
    Columns.BackColor = vbWhite
    ColumnsLabel.ForeColor = vbBlack
    If Columns.Value = "" Or _
      Not ValidColumnsName(Columns.Value, ColumnCount.Value) Then
        Columns.BackColor = vbRed
        ColumnsLabel.ForeColor = vbRed
    End If
End Sub

Private Sub CoulmnCountSpin_Change()
    ' when we increase or decrease the value
    ' update it in the field form
    ColumnCount.Value = CoulmnCountSpin.Value
End Sub


' Creating a table with a name, the number of columns
' specified and their columns
Private Sub Create_Click()
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    ColumnCountLabel.ForeColor = vbBlack
    ColumnCount.BackColor = vbWhite
    Columns.BackColor = vbWhite
    ColumnsLabel.ForeColor = vbBlack
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
    ColumnCount.Value = CoulmnCountSpin.Value
End Sub
