VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueriesTableWindow 
   Caption         =   "QueriesTable"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
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

Private Sub Querie_Click()
    Dim aux As Boolean
    
    aux = True
    
    If SelectorInput.Value = "" Or IsNumeric(SelectorInput.Value) Then
        aux = False
        errorOut ("Invalid table name")
    End If
    
    If aux = True Then
        
        For i = 1 To ActiveWorkbook.Worksheets.Count
            If SelectorInput.Value = ActiveWorkbook.Worksheets(i).name Then
                Set currentSheet = ActiveWorkbook.Worksheets(i)
               
               ' Daca prima celula din tabela este goala nu are rost sa cautam
               ' deoarece tabela e goala
                If currentSheet.Cells(1, "A").Value = "" Then
                    emptyRow = 1
                    MsgBox ("Tabela este goala! Nu a fost fasit nici un rezultat!")
                End If
                
                'To do parcurgere pt select prima linie din coloana
            
            End If
        Next i
        
        
        
    End If
    
    
End Sub

Private Sub SelectorInput_Change()

End Sub
