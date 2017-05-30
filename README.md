# VBA20


> VBA-20. Sa se implementeze un sistem care simuleaza lucrul cu o baza de date. Pentru fiecare tabel, datele vor fi retinute intr-o foaie de lucru separata. In fiecare asemenea foaie de lucru, pe prima linie se gasesc numele coloanelor din tabelul respectiv; numele foii de lucru este si numele tabelului.

##### Functionalitate:


* Crearea unui tabel, cu indicarea de catre utilizator a numelor campurilor.
* Stergerea unui tabel, ales dintr-o lista a tabelelor existente.
* Inserarea unei linii intr-un tabel, cu indicarea valorilor campurilor.
* Realizarea de interogari asupra tabelelor:
        O interogare poate returna unul sau mai multe campuri deja existente dintr-un tabel, dar nu si expresii mai complexe.
        La precizarea clauzei WHERE, conditiile sunt operatorii relationali standard pentru numere si siruri de caractere (==, !=, <, >, <=, >=), aplicati fie intre campuri, fie intre campuri si valori constante.
        Pot fi indicate conditii logice multiple, legate prin AND sau OR, dar nu se pot folosi ambele functii booleene in aceeasi interogare.
        In clauza WHERE se poate face referire si la campuri din alte tabele. Utilizatorul va putea alege din liste care contin tabelele existente si respectiv campurile tabelului indicat. 
* Stergerea uneia sau mai multor linii dintr-un tabel. Clauza WHERE se supune acelorasi conditii ca in cazul interogarilor.



Sistemul care simuleaza o baza de date implementat dispune de o interfata. Ca punct de intrare in interfata va avea un button localizat in foaia dba_start.
Acest buton va lansa menium principal, prin care utilizatorul poate executa diferite operatiuni asupra Workbook-ului curent.

```vbnet
Private Sub CancelMenu_Click()
	Unload Me
End Sub

Private Sub CreateTable_Click()
	CreateTableWindow.Show
End Sub

Private Sub DeleteTable_Click()
	DeleteTableWindow.Show
End Sub

Private Sub InsertTable_Click()
	InsertTableWindow.Show
End Sub

Private Sub QueriesTable_Click()
	QueriesTableWindow.Show
End Sub

```


Fiecare click in meniu va lansa cate o noua fereastra, care va da acces utilizatorului sa lanseze diferite operatiuni. Butonul Cancel are rol de a inchide meniul prin ```Unload Me``` care in Basic are rol de a dealoca memoria si de a inchide fereastra curenta.


Fiecare camp in aplicatia noastra detecteaza in timp real validitatea campurilor. Daca campul este invalid, labelul si campul va avea culoarea rosie. Asta se poate realiza prin bindingul la sistemul de event cu numele de ```*_Change```. Exemplu, validarea campului ```TableName```

```vbnet

' On every change in the TableName field
' check if we are dealing with valid characters or not
Private Sub TableName_Change()
    TableNameLabel.ForeColor = vbBlack
    TableName.BackColor = vbWhite
    
    If TableName.Value = "" Or _
    IsNumeric(TableName.Value) Then
        TableNameLabel.ForeColor = vbRed
        TableName.BackColor = vbRed
    End If
End Sub

```


# Creare tabel

```vbnet

' Create a new table in our database
' In our case a table is just a new sheet
Private Sub Create_Click()
	' create table validity checks
	 CreateTable TableName.Value, columnNames
End Sub

```


> Pentru fiecare tabel, datele vor fi retinute intr-o foaie de lucru separata.

Prin apelul acestei proceduri, vom crea tabelul intr-un worksheet nou si ca primul row vom stoca capul tabelului. Mai jos un snippet din functia de creare a tabelului.

```vbnet

' createTable
' @name - name of the table
'
' If the @name is empty this will be no-op.
Public Sub CreateTable(ByVal name As String, ByRef columnNames() As String)
  	'create table code checks here
    
    Dim i As Integer
    Dim n As Integer
    n = ArraySize(columnNames)
    For i = 1 To n
        With Cells(1, i)
            .Value = columnNames(i - 1)
            .Font.size = 14
            '.Width = 14
            .Interior.Color = RGB(188, 188, 188)
        End With
    Next i
End Sub

```


# Stergere tabel

> Stergerea unui tabel, ales dintr-o lista a tabelelor existente.

Prin selectarea din lista tabelelor curente si prin apasarea butonului de delete se va sterge tabelul(worksheet-ul) cat si alte date importante care s-au scris atunci cand tabelul s-a creat.


```vbnet

Private Sub Delete_Click()
    'validity code check'
    
    Dim err As Boolean
    err = DeleteTable(Me.ComboBoxName.Value)
    If err = False Then
        errorOut ("Cannot delete table, table does not exist")
        Exit Sub
    End If
    
    
    ' find raw that belongs to the table we want to delete
    ' and be sure to delete the types also
    Dim dba_start As Worksheet
    Set dba_start = ActiveWorkbook.Worksheets("dba_start")
    Dim rw As Range
    For Each rw In dba_start.Rows
        If dba_start.Cells(rw.Row, 1).Value = Me.ComboBoxName.Value Then
            dba_start.Cells(rw.Row, 1).EntireRow.Delete
            Exit For
        End If
    Next rw
    
    Me.Hide
    Unload Me
End Sub
```




La apelarea functiei ```DeleteTable```, atunci cand tabelul se va sterge, normal in excel va aparea un pop up cu intrebarea daca cu adevarat vrem sa stergem worksheet-ul, dar cum nu vrem sa apara mereu acel pop up, vom da disable la astfel de event-uri prin liniile

```vbnet
Application.DisplayAlerts = False ' disable events'
Application.DisplayAlerts = True ' re enable events'
```


# Inserarea datelor

> Inserarea unei linii intr-un tabel, cu indicarea valorilor campurilor.

Inserarea nu este valida atunci cand vrem sa inseram o valoare de un type diferit fata de cel declarat a coloanei sub care vrem sa inseram.
Fiecare type a fiecarei coloane dinttr-un tabel sunt memorate in worksheet-ul dba_start atunci cand tabelul este creat.

```vbnet

' InsertTable
'
' @list - list of values to be inserted
' @table - the table in which the values will be inserted
'
' If the list values does not match the column count of the tables this will return False
' If any operation will fail this will return false
' If the values are inserted and everything went fine, this will return True
Public Function InsertTable(ByVal TableName As String, ByVal list As String) As Boolean
'code
End Function

```

Fiecare inserarea va verifica si valida type-ul daca este la fel cu cel declarat pentru coloana respectiva. Daca valorile nu coincid atunci vom informa userul de *type mismatch*.

Parcurgem vba_start si extragem type-urile in ordinea gasit a coloanelor

```vbnet
  For Each rw In dba_start.Rows
        ' if we found the raw for the @TableName
        ' we need to take all types
        If dba_start.Cells(rw.Row, 1).Value = TableName Then
            ' for every non empty cell
            Do While dba_start.Cells(rw.Row, j) <> ""
                ReDim Preserve columnTypes(0 To ncolumns)
                columnTypes(ncolumns) = dba_start.Cells(rw.Row, j).Value ' get the type
                j = j + 1
                ncolumns = ncolumns + 1
            Loop
            ' after we've done with the type extraction just exit
            Exit For
        End If
Next rw

```


Toate valorile se vor da despartite prin token-ul ```,```. Pentru fiecare cell in parte insertam dupa ultim-ul row non-empty pe care o gasim
``` emptyRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row + 1 ```.
Atata timp cat type-urile coincid vom insera ``` currentSheet.Cells(emptyRow, j + 1).Value = arr(j)```. Daca type-urile nu coincid, vom iesi din functie si vom afisa catre user un warning pop up cu mesajul "Column values type mismatch".


```vbnet

    Dim j As Integer
    Dim currentSheet As Worksheet
    ' insertion table logic
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If TableName = ActiveWorkbook.Worksheets(i).name Then
            Set currentSheet = ActiveWorkbook.Worksheets(i)
                
            Dim emptyRow As Long
            emptyRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row + 1
                
            If currentSheet.Cells(1, "A").Value = "" Then
                emptyRow = 1
            End If
                
            Dim emptyCol As Long
            For j = 0 To size - 1
                If columnTypes(j) = "Integer" And IsNumeric(arr(j)) Then
                    currentSheet.Cells(emptyRow, j + 1).Value = arr(j)
                ElseIf columnTypes(j) = "String" And Not IsNumeric(arr(j)) Then
                     currentSheet.Cells(emptyRow, j + 1).Value = arr(j)
                Else
                    errorOut ("Column value type mismatch")
                    InsertTable = False
                    Exit Function
                End If
            Next j
        End If
    Next i

```