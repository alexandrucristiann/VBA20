VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueriesTableWindow 
   Caption         =   "QueriesTable"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14070
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
