Attribute VB_Name = "toolbar_grid"
Const toolbarname As String = "‘»ﬂ«  »‰Êœ"

Sub grid_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
  
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set grid_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With grid_button
         .DescriptionText = "≈œ—«Ã ‘»ﬂ… »‰Êœ"
          'Tooltip text when mouse if placed over button

         .Caption = "≈œ—«Ã ‘»ﬂ… »‰Êœ"
         'Text if Text in Icon is chosen

         .OnAction = "create_grid_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 3620
         
    End With
    
           
    Set vertical_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With vertical_button
         .DescriptionText = " ‰”Ìﬁ ⁄„Êœ „‰ «·»‰Êœ"
          'Tooltip text when mouse if placed over button

         .Caption = " ‰”Ìﬁ ⁄„Êœ „‰ «·»‰Êœ"
         'Text if Text in Icon is chosen

         .OnAction = "vertical_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 9979
         
    End With
    
    Set horizontal_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With horizontal_button
         .DescriptionText = " ‰”Ìﬁ ’› „‰ «·»‰Êœ"
          'Tooltip text when mouse if placed over button

         .Caption = " ‰”Ìﬁ ’› „‰ «·»‰Êœ"
         'Text if Text in Icon is chosen

         .OnAction = "horizontal_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

         .Picture = LoadPicture(get_root_dir & "toolbars\hor.jpg")
         
    End With
    
    oToolbar.Visible = True
End Sub


Sub create_grid_callback()
    On Error GoTo showerror
        make_grid_size
    GoTo fin
showerror:
        show_error " √ﬂœ „‰  ÕœÌœ »‰œ ·Ì „ «” »œ«·Â"
fin:
End Sub


Sub horizontal_callback()
    On Error GoTo showerror
        Dim max_height As Double
        If ActiveWindow.Selection.shaperange.Count > 1 Then
            max_height = 0
            For Each oshp In ActiveWindow.Selection.shaperange
                max_height = max(oshp.Height, max_height)
            Next oshp
            
            ActiveWindow.Selection.shaperange.Height = max_height
            ActiveWindow.Selection.shaperange.Align msoAlignTops, False
            
            If ActiveWindow.Selection.shaperange.Count > 2 Then
                ActiveWindow.Selection.shaperange.Distribute msoDistributeHorizontally, False
            End If

        End If
    
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub vertical_callback()
    On Error GoTo showerror
        Dim max_width As Double
        If ActiveWindow.Selection.shaperange.Count > 1 Then
            max_width = 0
            For Each oshp In ActiveWindow.Selection.shaperange
                max_width = max(oshp.Width, max_width)
            Next oshp
            
            ActiveWindow.Selection.shaperange.Width = max_width
            ActiveWindow.Selection.shaperange.Align msoAlignCenters, False
            
            If ActiveWindow.Selection.shaperange.Count > 2 Then
                ActiveWindow.Selection.shaperange.Distribute msoDistributeVertically, False
            End If

        End If
    
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub







