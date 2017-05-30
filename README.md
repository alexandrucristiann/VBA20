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