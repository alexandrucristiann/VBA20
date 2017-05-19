Attribute VB_Name = "base"
Public Const MAX_COLUMNS As Byte = 200
Public Const MIN_COLUMNS As Byte = 1

' ValidColumnName
' Checks when ever the number of column names from @list matches the @limit passed
' if the @list is empty this will return False
' if the limit is not in this interval of [MIN_COLUMNS,MAX_COLUMNS] it will return False
' TODO(hoenir) fix this.
Public Function ValidColumnsName(list As String, limit As Integer) As Boolean
    ValidColumnsName = True
    
    Dim arr() As String
    If list = "" Or _
        limit > MAX_COLUMNS Or _
        limit < MIN_COLUMNS Then
        ValidColumnsName = False
        Exit Function
    End If
     
    arr = Split(list, ",")
    
    If arr.Length <> limit Then
        ValidColumnsName = False
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To limit
        If arr(i) = "" Then
          ValidColumnsName = False
          Exit Function
        End If
    Next
    
End Function

