VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} animation 
   Caption         =   "„œÏ Õ—ﬂ«  «·‘—«∆Õ"
   ClientHeight    =   1590
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2244
   OleObjectBlob   =   "animation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Proceed_Click()
    On Error GoTo errorgui:
    
    Dim Sstart As Integer
    Dim send As Integer
    
    Sstart = Int(st.text)
    send = Int(ed.text)
    If send < Sstart Then
        GoTo errorgui
    End If
    
    hide
    primary_animation Sstart, send
    GoTo escape
errorgui:
    show_error "—«Ã⁄ «·œŒ·"
escape:
    
End Sub
