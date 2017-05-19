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
Public Function validateColumns(ByVal columns As String, ByVal limit As Byte) As Boolean
    Dim arr() As String
    Dim size As Byte
    validateColumns = True
    
    If columns = "" Or limit = 0 Then
        validateColumns = False
        Exit Function
    End If
    
    arr = Split(columns, ",", limit)
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

' error

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





