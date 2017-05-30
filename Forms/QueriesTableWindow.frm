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

Private Sub Frame1_Click()

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

Private Function getDelim(ByVal sir As String) As String
    
    If InStr(1, sir, "==") > 0 Then
        getDelim = "=="
    End If
    
    If InStr(1, sir, "!=") > 0 Then
        getDelim = "!="
    End If
    
    If InStr(1, sir, ">=") > 0 Then
        getDelim = ">="
    End If
    
    If InStr(1, sir, "<=") > 0 Then
        getDelim = "<="
    End If
    
    If InStr(1, sir, "<") > 0 Then
        getDelim = "<"
    End If
    
    If InStr(1, sir, ">") > 0 Then
        getDelim = ">"
    End If
    
End Function

Private Function computeByDelim(ByVal sir1 As String, ByVal sir2 As String, ByVal delim As String) As Boolean

    If delim = "==" Then
        computeByDelim = (sir1 = sir2)
    End If
    
    If delim = "!=" Then
        computeByDelim = (sir1 <> sir2)
    End If
    
    If delim = "<=" Then
        computeByDelim = (sir1 <= sir2)
    End If
    
    If delim = ">=" Then
        computeByDelim = (sir1 >= sir2)
    End If
    
    If delim = "<" Then
        computeByDelim = (sir1 < sir2)
    End If
    
    If delim = ">" Then
        computeByDelim = (sir1 > sir2)
    End If

End Function

Private Function isInArray(ByRef arr() As Integer, ByVal x As Integer) As Boolean
    sizeOfArray = UBound(arr, 1) - LBound(arr, 1)
    
    isInArray = False
    For i = 0 To sizeOfArray
        If arr(i) = x Then
            isInArray = True
        End If
    Next i
    
End Function

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
    
    rulesValues = Split(WhereInput.Value, "AND", 1000)
    rulesArraySize = UBound(rulesValues) - LBound(rulesValues) + 1
    
    Dim isAnd As Boolean
    Dim isOr As Boolean
    
    If rulesArraySize = 1 Then
        rulesValues = Split(WhereInput.Value, "OR", 1000)
        rulesArraySize = UBound(rulesValues) - LBound(rulesValues) + 1
        If rulesArraySize > 1 Then
            isOr = True
        End If
    Else
        isAnd = True
    End If
    
    InputSelectValue.Value = Replace(InputSelectValue.Value, " ", "")
    selectedValues = Split(InputSelectValue.Value, ",", 1000)
    size = UBound(selectedValues) - LBound(selectedValues) + 1
    
    Dim result(1000) As Integer
    Dim deScos(1000) As Integer
    
    Dim resultLen As Integer
    Dim deScosLen As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If DropdownFrom.Value = ActiveWorkbook.Worksheets(i).name Then
            Set currentSheet = ActiveWorkbook.Worksheets(i)
        End If
    Next i
    
    emptyCell = 1
    Do While currentSheet.Cells(1, emptyCell).Value <> ""
        emptyCell = emptyCell + 1
    Loop
    
    For i = 1 To emptyCell - 1
        dict.Add currentSheet.Cells(1, i).Value, i
    Next i
    
    
    resultLen = 0
    deScosLen = 0
    
    If aux = True Then
        
        For i = 1 To ActiveWorkbook.Worksheets.Count
            If DropdownFrom.Value = ActiveWorkbook.Worksheets(i).name Then
                Set currentSheet = ActiveWorkbook.Worksheets(i)
               
               ' Daca prima celula din tabela este goala nu are rost sa cautam
                
                emptyRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row + 1
               
                If currentSheet.Cells(2, "A").Value = "" Then
                    emptyRow = 2
                    MsgBox ("Tabela este goala! Nu a fost fasit nici un rezultat!")
                End If
                
                'Fac pentru mai multe coloane folosd virgula la Select
                
                For k = 0 To rulesArraySize - 1
                    
                    Dim auxArray() As String
                    'Split pe clauza where la atribuire pe A=A
                    Dim delimitator As String
                    
                    delimitator = getDelim(rulesValues(k))
                    
                    rulesValues(k) = Replace(rulesValues(k), " ", "")
                    auxArray = Split(rulesValues(k), delimitator, 1000)
                    
                    emptyCell = 1
                    
                    Do While currentSheet.Cells(1, emptyCell).Value <> ""
                        emptyCell = emptyCell + 1
                    Loop
                    
                    For ii = 1 To emptyCell - 1
                        
                        If currentSheet.Cells(1, ii).Value = auxArray(0) Then
                            
                            auxArray(0) = ii
                            
                        End If
                        
                    Next ii
                    
                    
                    For l = 2 To emptyRow - 1
                        'Merg pana la gasesc celula goala
                        'Debug -- MsgBox (currentSheet.Cells(l, auxArray(0)).Value)
                        
                        Dim toBeComparedWith As String
                        If InStr(1, auxArray(1), "'") > 0 Then
                            toBeComparedWith = Replace(auxArray(1), "'", "")
                        Else
                            toBeComparedWith = currentSheet.Cells(l, auxArray(1)).Value
                        End If
                        
                        If computeByDelim(currentSheet.Cells(l, CInt(auxArray(0))).Value, toBeComparedWith, delimitator) Then
                            
                            
                            If Not isInArray(result, l) Then
                                result(resultLen) = l
                                resultLen = resultLen + 1
                            End If
                            
                            
                            Else
                                If isInArray(result, l) = True And isAnd = True Then
                                    
                                    deScos(deScosLen) = l
                                    deScosLen = deScosLen + 1
                                
                                End If
                                
                        End If
                    Next l
                    
                Next k
                
            End If
        Next i
        
        
        For l = 0 To deScosLen - 1
            For ii = 0 To resultLen - 1
                If result(ii) = deScos(l) Then
                    For jj = ii + 1 To resultLen - 1
                        result(jj - 1) = result(jj)
                    Next jj
                    
                End If
            Next ii
            resultLen = resultLen - 1
        Next l
        
        
        
        Dim resultString As String
        Dim sizeOfLineArray As Integer
        
        For i = 0 To resultLen - 1
            
            resultString = ""
            
            For j = 0 To size - 1
                resultString = resultString & currentSheet.Cells(result(i), dict(selectedValues(j))).Value
                resultString = resultString & " "
            Next j
            
            MsgBox (resultString)
            
        Next i
    
        
    End If
    
    
End Sub
