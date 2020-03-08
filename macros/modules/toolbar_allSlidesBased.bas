Attribute VB_Name = "toolbar_allSlidesBased"
Const toolbarname As String = "⁄„·Ì«  «·⁄—Ê÷"

Sub slidebased_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set animate_primary_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With animate_primary_button
         .DescriptionText = "Õ—ﬂ«   √À—Ì… ··⁄—÷ «·—∆Ì”Ï(” «∆—)"
          'Tooltip text when mouse if placed over button

         .Caption = "Õ—ﬂ«   √À—Ì… ··⁄—÷ «·—∆Ì”Ï(” «∆—)"
         'Text if Text in Icon is chosen

         .OnAction = "animate_primary_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 346

    End With
    
    
    Set animate_secondary_button = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With animate_secondary_button

         .DescriptionText = "«“«·… Ã„Ì⁄ «·Õ—ﬂ«  «· √À—Ì…"
         
          'Tooltip text when mouse if placed over button

         .Caption = "«“«·… Ã„Ì⁄ «·Õ—ﬂ«  «· √À—Ì…"
         'Text if Text in Icon is chosen

         .OnAction = "animate_secondary_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 348
    End With
    
    
    Set remove_internal_button = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With remove_internal_button

         .DescriptionText = " ›—Ì€ «·Ê«‰"
         
          'Tooltip text when mouse if placed over button

         .Caption = " ›—Ì€ «·Ê«‰"
         'Text if Text in Icon is chosen

         .OnAction = "remove_internal_callback"

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 5872
    End With
    
    
    Set remove_internal_BW_button = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With remove_internal_BW_button

         .DescriptionText = " ›—Ì€ «»Ì÷ Ê«”Êœ"
         
          'Tooltip text when mouse if placed over button

         .Caption = " ›—Ì€ «»Ì÷ Ê«”Êœ"
         'Text if Text in Icon is chosen

         .OnAction = "remove_internal_BW_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 5876
    End With
    
    
    Set corners_button = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With corners_button

         .DescriptionText = " ⁄œÌ· «·≈ÿ«—"
         
          'Tooltip text when mouse if placed over button

         .Caption = " ⁄œÌ· «·≈ÿ«—"
         'Text if Text in Icon is chosen

         .OnAction = "corners_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 6781
    End With
    
    
    'Set highlighter_button = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    'With highlighter_button

         '.DescriptionText = "This is my first button"
          'Tooltip text when mouse if placed over button

         '.Caption = "Do Button1 Stuff"
         'Text if Text in Icon is chosen

         '.OnAction = "highlighter_callback"
         

         '.Style = msoButtonIcon
          ' Button displays as icon, not text or both

          '.FaceId = 6728
    'End With
    
    
    oToolbar.Visible = True
End Sub


Sub highlighter_callback()
    MsgBox "highlighter_callback"
End Sub

Function fix_all_bugs()
    MsgBox "”ÌﬁÊ„ «·»—‰«„Ã »„⁄«·Ã… «·Œÿ√ Ê≈⁄«œ… «·„Õ«Ê·… Ì—ÃÏ „—«Ã⁄… »⁄œ «·≈‰ Â«¡", Title:="„Õ«Ê·… «·≈’·«Õ"
    
    make_progressor
    slide_count = ActivePresentation.Slides.Count
    
    With ActivePresentation
        For Each osld In .Slides
            osld.shapes.Range.Cut
            osld.shapes.Paste
    
            DoEvents
            update_progressor Int(osld.SlideIndex / slide_count * 100)
    
        Next osld
    End With
    
    done_progressor
   

    'correction = False
    'While Not correction
    '    correction = True
    '    For Each shape In shapes
    '        If Not is_correct_shape(shape) Then
    '            shape.Cut
    '            slide.shapes.Paste
    '            correction = False
    '        End If
    '    Next shape
    'Wend
    
End Function

Sub corners_callback()
    On Error GoTo showerror
        cleanrs
        GoTo fin
showerror:
   show_error "ÕœÀ Œÿ√ Ê”ÌﬁÊ„ «·»—‰«„Ã »≈’·«Õ «·‘—«∆Õ° ≈–«  ﬂ—— ŸÂÊ— Â–« «·Œÿ√ Ì—ÃÏ ≈€·«ﬁ ﬂ· «·‰Ê«›– À„ ≈⁄«œ… › Õ «·»—‰«„Ã"
   fix_all_bugs
   cleanrs
   MsgBox " „ «·≈‰ Â«¡ „‰ ⁄„·Ì… «·≈’·«Õ", Title:="⁄„·Ì… «·≈’·«Õ"
  
fin:

End Sub


Sub animate_primary_callback()
    On Error GoTo showerror
        make_animation_range
        GoTo fin
showerror:
   show_error "ÕœÀ Œÿ√ Ê”ÌﬁÊ„ «·»—‰«„Ã »≈’·«Õ «·‘—«∆Õ° ≈–«  ﬂ—— ŸÂÊ— Â–« «·Œÿ√ Ì—ÃÏ ≈€·«ﬁ ﬂ· «·‰Ê«›– À„ ≈⁄«œ… › Õ «·»—‰«„Ã"
   fix_all_bugs
   make_animation_range
fin:
End Sub

Sub animate_secondary_callback()
    On Error GoTo showerror
        make_un_animation_range
        GoTo fin
showerror:
   show_error "ÕœÀ Œÿ√ Ê”ÌﬁÊ„ «·»—‰«„Ã »≈’·«Õ «·‘—«∆Õ° ≈–«  ﬂ—— ŸÂÊ— Â–« «·Œÿ√ Ì—ÃÏ ≈€·«ﬁ ﬂ· «·‰Ê«›– À„ ≈⁄«œ… › Õ «·»—‰«„Ã"
   fix_all_bugs
   make_un_animation_range
fin:
End Sub

Sub remove_internal_callback()
    MsgBox "”ÌﬁÊ„ «·»—‰«„Ã » ›—Ì€ «·⁄—÷ ··ÿ»«⁄… «·√·Ê«‰ Ì—ÃÏ „—«Ã⁄… «·√·Ê«‰ €Ì— «·Ê«÷Õ…", Title:="⁄„·Ì… «· ›—Ì€ √·Ê«‰"
  
    On Error GoTo showerror
        removing_interior_foreground_color
        GoTo fin
showerror:
   show_error "ÕœÀ Œÿ√ Ê”ÌﬁÊ„ «·»—‰«„Ã »≈’·«Õ «·‘—«∆Õ° ≈–«  ﬂ—— ŸÂÊ— Â–« «·Œÿ√ Ì—ÃÏ ≈€·«ﬁ ﬂ· «·‰Ê«›– À„ ≈⁄«œ… › Õ «·»—‰«„Ã"
   fix_all_bugs
   removing_interior_foreground_color
fin:
End Sub

Sub remove_internal_BW_callback()
    On Error GoTo showerror
        removing_interior_foreground_color_black_white
        GoTo fin
showerror:
   show_error "ÕœÀ Œÿ√ Ê”ÌﬁÊ„ «·»—‰«„Ã »≈’·«Õ «·‘—«∆Õ° ≈–«  ﬂ—— ŸÂÊ— Â–« «·Œÿ√ Ì—ÃÏ ≈€·«ﬁ ﬂ· «·‰Ê«›– À„ ≈⁄«œ… › Õ «·»—‰«„Ã"
   fix_all_bugs
   removing_interior_foreground_color_black_white
fin:
End Sub


Sub slidebased_enable()
    enable_bar toolbarname
End Sub

Sub slidebased_disable()
    disble_bar toolbarname
End Sub

'.DescriptionText = " ⁄œÌ· «·≈ÿ«—"
'          .DescriptionText = " ›—Ì€ «»Ì÷ Ê«”Êœ"
'    .DescriptionText = " ›—Ì€ «·Ê«‰"
'   .DescriptionText = "«“«·… Ã„Ì⁄ «·Õ—ﬂ«  «· √À—Ì…"
' .DescriptionText = "Õ—ﬂ«   √À—Ì… ··⁄—÷ «·—∆Ì”Ï(” «∆—)"



