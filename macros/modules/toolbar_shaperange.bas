Attribute VB_Name = "toolbar_shaperange"
Const toolbarname As String = "⁄„·Ì«  «·√‘ﬂ«·"

Sub shaperange_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set parenthesis_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With parenthesis_button
         .DescriptionText = " ·ÊÌ‰ „« »Ì‰ «·«ﬁÊ«”"
         .Caption = " ·ÊÌ‰ „« »Ì‰ «·«ﬁÊ«”"
         .OnAction = "parenthesis_callback"
         .Style = msoButtonIcon
         .FaceId = 2636
    End With
    
    
    Set numbers_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With numbers_button
         .DescriptionText = " ·ÊÌ‰ «·«—ﬁ«„ Ê  Ê«—ÌŒ"
         .Caption = " ·ÊÌ‰ «·«—ﬁ«„ Ê  Ê«—ÌŒ"
         .OnAction = "numbers_callback"
         .Style = msoButtonIcon
         .FaceId = 4207
    End With
    
        
    Set DistributeVertically_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With DistributeVertically_button
         .DescriptionText = " Ê“Ì⁄ ⁄„ÊœÏ"
         .Caption = " Ê“Ì⁄ ⁄„ÊœÏ"
         .OnAction = "DistributeVertically_callback"
         .Style = msoButtonIcon
          .FaceId = 3138
    End With
        
    
    
    Set AlignTops_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With AlignTops_button
         .DescriptionText = "„Õ«–«… ≈·Ï √⁄·Ï"
         .Caption = "„Õ«–«… ≈·Ï √⁄·Ï"
         .OnAction = "AlignTops_callback"
         .Style = msoButtonIcon
         .FaceId = 3072
    End With
    
    
    Set AlignLefts_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With AlignLefts_button
         .DescriptionText = "„Õ«–«… ≈·Ï «·Ì”«—"
         .Caption = "„Õ«–«… ≈·Ï «·Ì”«—"
         .OnAction = "AlignLefts_callback"
         .Style = msoButtonIcon
         .FaceId = 3070
    End With
    
    Set AlignBottoms_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With AlignBottoms_button
         .DescriptionText = "„Õ«–«… ≈·Ï «·√”›·"
         .Caption = "„Õ«–«… ≈·Ï «·√”›·"
         .OnAction = "AlignBottoms_callback"
         .Style = msoButtonIcon
         .FaceId = 3073
    End With


    Set AlignCenters_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With AlignCenters_button
         .DescriptionText = "„Õ«–«… ≈·Ï «·Ê”ÿ"
         .Caption = "„Õ«–«… ≈·Ï «·Ê”ÿ"
         .OnAction = "AlignCenters_callback"
         .Style = msoButtonIcon
         .FaceId = 3074
    End With

    Set AlignRights_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With AlignRights_button
         .DescriptionText = "„Õ«–«… ≈·Ï «·Ì„Ì‰"
         .Caption = "„Õ«–«… ≈·Ï «·Ì„Ì‰"
         .OnAction = "AlignRights_callback"
         .Style = msoButtonIcon
         .FaceId = 3071
    End With


    Set borders_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With borders_button
         .DescriptionText = "ÕœÊœ «·Ãœ«Ê· Ê«·’Ê—"
         .Caption = "ÕœÊœ «·Ãœ«Ê· Ê«·’Ê—"
         .OnAction = "borders_callback"
         .Style = msoButtonIcon
         .Picture = LoadPicture(get_root_dir & "toolbars\borders.jpg")
    End With
    
    Set increaseLineSpace_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With increaseLineSpace_button
         .DescriptionText = "“Ì«œ…  »«⁄œ «·√”ÿ—"
         .Caption = "“Ì«œ…  »«⁄œ «·√”ÿ—"
         .OnAction = "increaseLineSpace_callback"
         .Style = msoButtonIcon
         .FaceId = 698
    End With
    
    Set decreaseLineSpace_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With decreaseLineSpace_button
         .DescriptionText = " ﬁ·Ì·  »«⁄œ «·√”ÿ—"
         .Caption = " ﬁ·Ì·  »«⁄œ «·√”ÿ—"
         .OnAction = "decreaseLineSpace_callback"
         .Style = msoButtonIcon
         .FaceId = 699
    End With
    
    
    Set squeeze_height_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With squeeze_height_button
         .DescriptionText = "÷€ÿ ﬁÌ«”Ï"
         .Caption = "÷€ÿ ﬁÌ«”Ï"
         .OnAction = "squeeze_height_callback"
         .Style = msoButtonIcon
         .FaceId = 3120
    End With
    
    
    ' add hidden font button
    Set reformat_button = oToolbar.Controls.Add(Type:=msoControlButton)
    With reformat_button
         .DescriptionText = "„œ «Œ— ﬂ·„…"
         .Caption = "„œ «Œ— ﬂ·„…"
         .OnAction = "hidden_font"
         .Style = msoButtonIcon
         .FaceId = 123
    End With
    
    
    oToolbar.Visible = True
    inner_colors_disable
End Sub


Sub parenthesis_callback()
    On Error GoTo showerror
        'shaperange_correction
        parenthesized_text_colorization_shprng ActiveWindow.Selection.shaperange
        GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub numbers_callback()
    On Error GoTo showerror
        'shaperange_correction
        Set regX = get_numbers_and_dates_regex()
        coloring_numbers_and_dates_shprng ActiveWindow.Selection.shaperange, regX
        GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub borders_callback()
    On Error GoTo showerror
        For Each shape In get_linear_shapes(ActiveWindow.Selection.shaperange)
            If get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        For i = 1 To 4
                            Cell.Borders.Item(i).ForeColor.rgb = get_table_border_rgb()
                            Cell.Borders.Item(i).Weight = 3
                        Next i
                    Next Cell
                Next Row
            ElseIf get_shape_type(shape) = "picture" Then
                With shape
                    .Line.ForeColor.rgb = get_image_border_rgb()
                    .Line.Weight = 2.25
                    .Line.Visible = msoTrue
                    .Line.Style = msoLineSingle
                End With
            End If
        Next shape
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub


Sub AlignRights_callback()
    On Error GoTo showerror
    
    If ActiveWindow.Selection.shaperange.Count > 1 Then
        ActiveWindow.Selection.shaperange.Align msoAlignRights, False
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub AlignCenters_callback()
    On Error GoTo showerror
        
    If ActiveWindow.Selection.shaperange.Count > 1 Then
        ActiveWindow.Selection.shaperange.Align msoAlignCenters, False
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub AlignBottoms_callback()
    On Error GoTo showerror
    
    
    If ActiveWindow.Selection.shaperange.Count > 1 Then
        ActiveWindow.Selection.shaperange.Align msoAlignBottoms, False
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub AlignLefts_callback()
    On Error GoTo showerror
    If ActiveWindow.Selection.shaperange.Count > 1 Then
        ActiveWindow.Selection.shaperange.Align msoAlignLefts, False
    End If
    
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub
    
  
Sub AlignTops_callback()
    On Error GoTo showerror
    If ActiveWindow.Selection.shaperange.Count > 1 Then
        ActiveWindow.Selection.shaperange.Align msoAlignTops, False
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub DistributeVertically_callback()
    On Error GoTo showerror
    
    If ActiveWindow.Selection.shaperange.Count > 2 Then
        ActiveWindow.Selection.shaperange.Distribute msoDistributeVertically, False
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

    

Sub increaseLineSpace_callback()
    On Error GoTo showerror
        For Each shape In get_linear_shapes(ActiveWindow.Selection.shaperange):
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.MarginBottom = 3.6
                shape.TextFrame.MarginLeft = 7.2
                shape.TextFrame.MarginRight = 7.2
                shape.TextFrame.MarginTop = 3.6
                
                shape.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                
                shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin + 0.1
                shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
                shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
            End If
        Next shape
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub


Sub decreaseLineSpace_callback()
    On Error GoTo showerror
        For Each shape In get_linear_shapes(ActiveWindow.Selection.shaperange):
            If get_shape_type(shape) = "textbox" Then
            
                shape.TextFrame.MarginBottom = 3.6
                shape.TextFrame.MarginLeft = 7.2
                shape.TextFrame.MarginRight = 7.2
                shape.TextFrame.MarginTop = 3.6
                
                shape.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                
                shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin - 0.1
                shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
                shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
            End If
        Next shape
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub squeeze_height_callback()
    On Error GoTo showerror
        squeeze_height_shprng ActiveWindow.Selection.shaperange
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:

End Sub



Sub hidden_font()
    On Error GoTo showerror
        reformat_text_shprng ActiveWindow.Selection.shaperange
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub


Sub shaperange_enable()
    enable_bar toolbarname
End Sub

Sub shaperange_disable()
    disble_bar toolbarname
End Sub



