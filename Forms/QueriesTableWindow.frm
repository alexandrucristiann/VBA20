VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueriesTableWindow 
   Caption         =   "QueriesTable"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12510
   OleObjectBlob   =   "QueriesTableWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QueriesTableWindow"
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

Private Sub DropdownFrom_Change()
    DropdownFrom.ForeColor = vbBlack
    DropdownFrom.BackColor = vbWhite
End Sub

Private Sub Userform_Initialize()
    For i = Me.DropdownFrom.ListCount - 1 To 0 Step -1
        Me.DropdownFrom.RemoveItem i
    Next i
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).name <> "dba_start" Then
            Me.DropdownFrom.AddItem ActiveWorkbook.Worksheets(i).name
        End If
    Next i
End Sub

Private Sub Querie_Click()
    Dim aux As Boolean
    
    aux = True
    
    If DropdownFrom.Value = "" Or IsNumeric(DropdownFrom.Value) Or DropdownFrom.Value = "Choose Table" Then
        aux = False
        'errorOut ("Invalid table name")
        DropdownFrom.Text = "Invalid table name! Select from dropdown!"
        DropdownFrom.ForeColor = vbWhite
        DropdownFrom.BackColor = vbRed
        
        
        
    End If
    
    
    
    Dim selectedValues() As String
    Dim size As Byte
    
    Dim rulesValues() As String
    Dim rulesArraySize As Byte
    
    rulesValues = Split(WhereInput.Value, ",", 1000)
    rulesArraySize = UBound(rulesValues) - LBound(rulesValues) + 1
    
    selectedValues = Split(InputSelectValue.Value, ",", 1000)
    size = UBound(selectedValues) - LBound(selectedValues) + 1
    
    Dim result As String
    
    If aux = True Then
        
        For i = 1 To ActiveWorkbook.Worksheets.Count
            If DropdownFrom.Value = ActiveWorkbook.Worksheets(i).name Then
                Set currentSheet = ActiveWorkbook.Worksheets(i)
               
               ' Daca prima celula din tabela este goala nu are rost sa cautam
                
                emptyRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row + 1
               
                If currentSheet.Cells(1, "A").Value = "" Then
                    emptyRow = 1
                    MsgBox ("Tabela este goala! Nu a fost fasit nici un rezultat!")
                End If
                
                'Fac pentru mai multe coloane folosd virgula la Select
                
                For k = 0 To rulesArraySize - 1
                    
                    Dim auxArray() As String
                    'Split pe clauza where la atribuire pe A=A
                    auxArray = Split(rulesValues(k), "=", 1000)
                    For l = 1 To emptyRow
                        'Merg pana la gasesc celula goala
                        'Debug -- MsgBox (currentSheet.Cells(l, auxArray(0)).Value)
                        If currentSheet.Cells(l, auxArray(0)).Value = auxArray(1) Then
                            result = ""
                            For j = 0 To size - 1
                                result = result & currentSheet.Cells(l, selectedValues(j)).Value
                                result = result & " "
                            Next j
                            'merge doar daca fac A=A nu si A = A
                            MsgBox (result)
                        End If
                    Next l
                    
                Next k
                
            End If
        Next i
        
        MsgBox (result)
        
    End If
    
    
End Sub

