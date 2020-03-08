VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grid_creator 
   Caption         =   "«‰‘«¡ ‘»ﬂ… »‰Êœ"
   ClientHeight    =   1530
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2244
   OleObjectBlob   =   "grid_creator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grid_creator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Proceed_Click()
    On Error GoTo errorgui:
    
    Dim rows As Integer
    Dim columns As Integer
    
    rows = Int(rows_tb.text)
    columns = Int(columns_tb.text)
    If send < Sstart Then
        GoTo errorgui
    End If
    
    hide
    create_grid_hidden rows, columns
    
    GoTo escape
errorgui:
    show_error "—«Ã⁄ «·œŒ·"
escape:
    
End Sub
