Attribute VB_Name = "toolbar_about"
Const toolbarname As String = "��.�.�.�"

Sub about_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
  
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set about_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With about_button
         .DescriptionText = "�� �������"
          'Tooltip text when mouse if placed over button

         .Caption = "�� �������"
         'Text if Text in Icon is chosen

         .OnAction = "show_about"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 59
         
    End With
    
    Set config_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With config_button
         .DescriptionText = "����� �������"
          'Tooltip text when mouse if placed over button

         .Caption = "����� �������"
         'Text if Text in Icon is chosen

         .OnAction = "open_notepad"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 1763
         
    End With
    
    
     Set config_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With config_button
         .DescriptionText = "����� �����"
          'Tooltip text when mouse if placed over button

         .Caption = "����� �����"
         'Text if Text in Icon is chosen

         .OnAction = "reload"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 6513
         
    End With
    

    Set remove_all_bars = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With remove_all_bars
         .DescriptionText = "����� �������"
          'Tooltip text when mouse if placed over button

         .Caption = "����� �������"
         'Text if Text in Icon is chosen

         .OnAction = "remove_all"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 3265
         
    End With
    
    
    
    oToolbar.Visible = True
End Sub

Sub remove_all()
    unloadForm.Show vbModeless
End Sub


Sub show_about()
 MsgBox _
        "        ����� ����� �������� ������ �������� ������� ������" & vbCrLf & _
        "        (����� ������/ ���� ���� ������ (��� ��� ���������" & vbCrLf & _
        "" & vbCrLf & _
        "                                      : ��� �����" & vbCrLf & _
        "      (���� / ���� ���� ��� (���� ��� ����� ��� ����� 2019" & vbCrLf & _
        "      (���� / ���� ����� ����� (���� ��� ����� �� ����� 2019" & vbCrLf & _
        "               (���� / ���� ���� ���� (���� ���� ������" & vbCrLf & _
        "                           ���� / ���� ����� �����" & vbCrLf _
        , Title:="                              ���� ����� ������ �������"

End Sub


Sub open_notepad()
    Shell "notepad.exe " & get_root_dir & "config.komy.txt", vbNormalFocus
End Sub



Sub reload()
    reloadForm.Show vbModeless
End Sub









