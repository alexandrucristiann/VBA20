Attribute VB_Name = "main"
Option Explicit

Public Sub main()
End Sub

' validateColumns
'
' @columns - list of comma separated strings
' @limit - the numeber of comma separated strings
'
' If the specified format is invalid this will return False
Public Function validateColumns(ByVal Columns As String, ByVal limit As Byte) As Boolean
    Dim arr() As String
    Dim size As Byte
    validateColumns = True
    
    If Columns = "" Or limit = 0 Then
        validateColumns = False
        Exit Function
    End If
    
    arr = Split(Columns, ",", limit)
    size = UBound(arr) - LBound(arr) + 1 ' compute the size of the array
    
    If size <> limit Or size = 0 Then
        validateColumns = False
        Exit Function
    End If
    
    ' In order to "for each" we need this type to be variant
    ' I know it's weird but heh...
    Dim a As Variant
    ' check for aditional blank elements
    For Each a In arr
        If a = "" Then
            validateColumns = False
            Exit Function
        End If
    Next
    
    'check if we have duplicates
    Dim i, j As Integer
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) = arr(j) Then ' we found duplicates
                validateColumns = False
                Exit Function
            End If
        Next j
    Next i
    
End Function

' errorOut

' @message - the error message
'
' Display a MsgBox in the error format specified by @message
' If the @message is empty this will be no-op
Public Sub errorOut(ByVal message As String)
    If message = "" Then
        Exit Sub
    End If
    
    MsgBox message, vbExclamation, "Application Error"
End Sub


' createTable
' @name - name of the table
'
' If the @name is empty this will ne no-op.
Public Sub CreateTable(ByVal name As String, ByVal n As Integer, ByRef columnsNames() As String)
    If name = "" Or n = 0 Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    With ThisWorkbook
        ' insert the new worksheet at the end of the worksheet list
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.name = name
    End With
    
    Sheets(name).Select

    Dim i As Integer
    For i = 0 To n
    ' TODO(hoenir) Fix this.
        Cells(0, i).Select
        ActiveCell.Value = columnsNames(0)
    Next i
End Sub

' deleteTable
' @name - name of the table
'
' If the @name is empty this will be no-op
' If the table(aka sheet) is found it will delete it and return True
' If the table is not found it will return False
Public Function DeleteTable(ByVal name As String) As Boolean
    If name = "" Then
        Exit Function
    End If
    
    Dim i As Integer
    Dim found As Boolean
    
    found = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).name = name Then
            found = True
            Exit For
        End If
    Next i
    
    
    If found = True Then
        ' delete Wroksheet without displaying messages
        Application.DisplayAlerts = False
        Worksheets(name).Delete
        Application.DisplayAlerts = True
    End If
    
    DeleteTable = found
End Function




