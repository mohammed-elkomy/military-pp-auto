VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} unloadForm 
   Caption         =   "����� �������"
   ClientHeight    =   1230
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3588
   OleObjectBlob   =   "unloadForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "unloadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    On Error GoTo escape
       CommandBars("��.�.�.�").Delete
       CommandBars("������ ������").Delete
       CommandBars("����� ����").Delete
       CommandBars("����� ������").Delete
       CommandBars("������ �������").Delete
       CommandBars("�����").Delete
       CommandBars("����� ������").Delete
       CommandBars("place").Delete
       CommandBars("place2").Delete
escape:
    
    hide
End Sub

