VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertTableWindow 
   Caption         =   "Insert Table"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9480
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
    Dim aux As Boolean
    
    aux = True
    
    If TableInsertName.Value = "" Or IsNumeric(TableInsertName.Value) Then
        aux = False
        errorOut ("Invalid table name")
    End If
    
    If ValueCount.Value = "" Then
        aux = False
        errorOut ("Invalid Value Count")
    End If
    
    Dim Valori() As String
    Dim size As Byte
    
    If Values = "" Then
        aux = False
    Else
        Valori = Split(Values.Value, ",", 1000)
        size = UBound(Valori) - LBound(Valori) + 1
        If size <> ValueCount.Value Then
            aux = False
            errorOut ("Invalid values")
        End If
    End If
    
    
    If aux = True Then
        
        For i = 1 To ActiveWorkbook.Worksheets.Count
            If TableInsertName.Value = ActiveWorkbook.Worksheets(i).name Then
                Set currentSheet = ActiveWorkbook.Worksheets(i)
                
                Dim emptyRow As Long
                emptyRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row + 1
                
                If currentSheet.Cells(1, "A").Value = "" Then
                    emptyRow = 1
                End If
                
                Dim emptyCol As Long
                For j = 0 To ValueCount.Value - 1
                    currentSheet.Cells(emptyRow, j + 1).Value = Valori(j)
                Next j
            
            End If
        Next i
        
        MsgBox ("Success!!!")
        
    End If
    
End Sub


Private Sub ValueCountSpin_Change()
    ValueCount.Value = ValueCountSpin.Value
End Sub

Private Sub TableInsertName_Change()
    TableInsertName.FontSize = 15
End Sub

Private Sub ValueCount_Change()
    ValueCount.FontSize = 15
End Sub

Private Sub Values_Change()
    Values.FontSize = 15
End Sub
