Attribute VB_Name = "toolbar_singleslide"
Const toolbarname As String = " ‰”Ìﬁ"


Sub optimizer_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
  
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set optimal_finder_button = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With optimal_finder_button
         .DescriptionText = " ‰”Ìﬁ ‘—ÌÕ… „⁄  ⁄œÌ· «·ŒÿÊÿ"
          'Tooltip text when mouse if placed over button

         .Caption = " ‰”Ìﬁ ‘—ÌÕ… „⁄   ⁄œÌ· «·ŒÿÊÿ"
         'Text if Text in Icon is chosen

         .OnAction = "optimal_finder_callback"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 509
         
    End With
    
    Set optimal_finder_button_without_reformating = oToolbar.Controls.Add(Type:=msoControlButton)
    ' And set some of the button's properties
    With optimal_finder_button_without_reformating
         .DescriptionText = " ‰”Ìﬁ ‘—ÌÕ… »œÊ‰  ⁄œÌ· «·ŒÿÊÿ"
          'Tooltip text when mouse if placed over button

         .Caption = " ‰”Ìﬁ ‘—ÌÕ… »œÊ‰   ⁄œÌ· «·ŒÿÊÿ"
         'Text if Text in Icon is chosen

         .OnAction = "optimal_finder_callback_without_reformating"
         

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

          .FaceId = 3051
         
    End With
    
    
    oToolbar.Visible = True
    
    'generate_heights_file
    heights = read_heights_file()

End Sub



Sub optimal_finder_callback()
is_copied = False
    On Error GoTo showerror

        With ActivePresentation
            .Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutText
            .Slides(ActivePresentation.Slides.Count).Layout = ppLayoutBlank
            .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Range.Copy
            .Slides(ActivePresentation.Slides.Count).shapes.Paste
        End With
        
        is_copied = True
        
        Dim heights() As Double
        heights = get_heights()
        
        find_optimal_line_spacing heights, ActiveWindow.View.slide
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
        
        If is_copied Then
            'an error occured cut from backup
            With ActivePresentation
                .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Range.Delete
                .Slides(ActivePresentation.Slides.Count).shapes.Range.Cut
                .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Paste
            End With
            
        End If
fin:
       If is_copied Then
            ActivePresentation.Slides(ActivePresentation.Slides.Count).Delete 'delete the backup
       End If
      
End Sub

Sub optimal_finder_callback_without_reformating()
is_copied = False
    On Error GoTo showerror
        With ActivePresentation
            .Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutText
            .Slides(ActivePresentation.Slides.Count).Layout = ppLayoutBlank
            .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Range.Copy
            .Slides(ActivePresentation.Slides.Count).shapes.Paste
        End With
        
        is_copied = True
        
        Dim heights() As Double
        heights = get_heights()
        find_optimal_line_spacing_without_reformating heights, ActiveWindow.View.slide
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
        
        If is_copied Then
            'an error occured cut from backup
            With ActivePresentation
                .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Range.Delete
                .Slides(ActivePresentation.Slides.Count).shapes.Range.Cut
                .Slides(ActiveWindow.View.slide.SlideIndex).shapes.Paste
            End With
            
        End If
fin:
       If is_copied Then
            ActivePresentation.Slides(ActivePresentation.Slides.Count).Delete 'delete the backup
       End If
        
    
Application.ActiveWindow.View.GotoSlide ActiveWindow.View.slide.SlideIndex
End Sub

Sub singleslide_enable()
    enable_bar toolbarname
End Sub

Sub singleslide_disable()
    disble_bar toolbarname
End Sub






