Attribute VB_Name = "toolbar_about"
Const toolbarname As String = "Â‹. .ﬁ.„"

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
         .DescriptionText = "⁄‰ «·≈÷«›…"
          'Tooltip text when mouse if placed over button

         .Caption = "⁄‰ «·≈÷«›…"
         'Text if Text in Icon is chosen

         .OnAction = "show_about"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 59
         
    End With
    
    Set config_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With config_button
         .DescriptionText = " ⁄œÌ· «·≈÷«›…"
          'Tooltip text when mouse if placed over button

         .Caption = " ⁄œÌ· «·≈÷«›…"
         'Text if Text in Icon is chosen

         .OnAction = "open_notepad"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 1763
         
    End With
    
    
     Set config_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With config_button
         .DescriptionText = "«⁄«œ…  Õ„Ì·"
          'Tooltip text when mouse if placed over button

         .Caption = "«⁄«œ…  Õ„Ì·"
         'Text if Text in Icon is chosen

         .OnAction = "reload"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 6513
         
    End With
    

    Set remove_all_bars = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With remove_all_bars
         .DescriptionText = "≈“«·… «·≈÷«›…"
          'Tooltip text when mouse if placed over button

         .Caption = "≈“«·… «·≈÷«›…"
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
        "        ≈÷«›… ÂÌ∆‹… «· œ—Ì‹» «·ﬁÊ«  «·„”·‹Õ… · ‰”Ì‹ﬁ «·⁄—Ê÷" & vbCrLf & _
        "        ( ’„Ì„ Ê ‰›Ì–/ „Õ„œ ⁄·«¡ «·ﬂÊ„Ï (›—⁄ ‰Ÿ„ «·„⁄·Ê„« " & vbCrLf & _
        "" & vbCrLf & _
        "                                      :  Õ  ≈‘—«›" & vbCrLf & _
        "      (⁄„Ìœ / Ê·Ìœ ”„Ì— –ﬂÏ (—∆Ì” ›—⁄ «·‰Ÿ„ Õ Ï ÌÊ·ÌÊ 2019" & vbCrLf & _
        "      (⁄ﬁÌœ / Ê·Ìœ ‰»Ì‹· ‘Êﬁ‹Ï (—∆Ì” ›—⁄ «·‰Ÿ„ „‰ ÌÊ·ÌÊ 2019" & vbCrLf & _
        "               („ﬁœ„ / ”«„Õ „Õ„œ Õ«›Ÿ (ﬁ«∆œ „—ﬂ“ «·Õ«”»" & vbCrLf & _
        "                           ‰ﬁÌ» / √Õ„œ „’ÿ›Ï «·”Ìœ" & vbCrLf _
        , Title:="                              ÂÌ∆…  œ—Ì» «·ﬁÊ«  «·„”·Õ…"

End Sub


Sub open_notepad()
    Shell "notepad.exe " & get_root_dir & "config.komy.txt", vbNormalFocus
End Sub



Sub reload()
    reloadForm.Show vbModeless
End Sub









