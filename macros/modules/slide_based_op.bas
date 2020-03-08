Attribute VB_Name = "slide_based_op"
Function get_hieght_sum(heights() As Double, line_spacing As Integer, shapes_no_lines As Variant, shapes_font_size As Variant, fixed_height As Double)
    sum = 0
    shape_index = 0
    For Each no_lines In shapes_no_lines
        sum = sum + heights(line_spacing, no_lines, shapes_font_size.Item(shape_index))
        shape_index = shape_index + 1
    Next no_lines
    get_hieght_sum = sum + fixed_height
End Function

Function get_padding_percentage(search_line_spacing As Integer, heights() As Double, available_vertical As Double, shapes_no_lines As Variant, shapes_font_size As Variant, fixed_height As Double)
    sum = get_hieght_sum(heights, search_line_spacing, shapes_no_lines, shapes_font_size, fixed_height)
    get_padding_percentage = (1 - sum / available_vertical) * 100
End Function

Function convex_cost_function_for_alignment(target_percentage As Double, search_line_spacing As Integer, heights() As Double, available_vertical As Double, shapes_no_lines As Variant, shapes_font_size As Variant, fixed_height As Double)
    ' this represent a convex cost function for alignmens optimized through ternary search
    padding_percentage = get_padding_percentage(search_line_spacing, heights, available_vertical, shapes_no_lines, shapes_font_size, fixed_height)
    convex_cost_function_for_alignment = Abs(padding_percentage - target_percentage)
End Function

Function find_optimal_line_spacing(heights() As Double, osld As Variant)
    find_optimal_line_spacing_hidden heights, osld, 0
End Function

Function find_optimal_line_spacing_without_reformating(heights() As Double, osld As Variant)
    find_optimal_line_spacing_hidden heights, osld, -1
End Function

Function get_title_width(title_shp As Variant)
    Dim title_left As Double
    Dim title_right As Double
    Dim diff1 As Double
    Dim diff2 As Double
    
    title_left = -10
    title_right = title_shp.Parent.CustomLayout.Width * 2
    
    For Each oshp In title_shp.Parent.Master.shapes
        If oshp.Left + oshp.Width < title_shp.Parent.CustomLayout.Width * 0.35 Then
            title_left = max(title_left, oshp.Left + oshp.Width)
        End If
        
        If oshp.Left + oshp.Width > title_shp.Parent.CustomLayout.Width * 0.65 Then
            title_right = min(title_right, oshp.Left)
        End If
    Next oshp
    diff1 = title_shp.Parent.CustomLayout.Width - title_right
    diff2 = title_left
    delta = max(diff1, diff2)
    title_shp.Width = title_shp.Parent.CustomLayout.Width - delta * 2.12
End Function


Function get_initial_top_position(osld As Variant)
    Dim initial_top As Double
    initial_top = osld.CustomLayout.Height * 2
    For Each oshp In osld.Master.shapes
        If oshp.Top + oshp.Height < osld.CustomLayout.Height * 0.35 Then
            initial_top = min(initial_top, oshp.Top + oshp.Height)
        End If
    Next oshp
    
    If initial_top > osld.CustomLayout.Height Then
        initial_top = osld.CustomLayout.Height * 0.055 - 10
    End If
    
    get_initial_top_position = initial_top + 10
End Function

Function get_initial_bottom_position(osld As Variant)
    Dim initial_bottom As Double
    initial_bottom = osld.CustomLayout.Height * 0.963 + 5
    For Each oshp In osld.Master.shapes
        If oshp.Top > osld.CustomLayout.Height * 0.65 Then
            initial_bottom = min(initial_bottom, oshp.Top)
        End If
    Next oshp

    get_initial_bottom_position = initial_bottom - 5
End Function

Function is_fixed(extended_shape As Variant)
    is_fixed = (get_shape_type(extended_shape.shape) <> "textbox") And (extended_shape.shape_bbox.area() / (ActiveWindow.View.slide.CustomLayout.Height * ActiveWindow.View.slide.CustomLayout.Width) * 100 < 5)
End Function

Function find_optimal_line_spacing_hidden(heights() As Double, osld As Variant, depth As Integer)
    ' cleaning
    adjust_corners_slide osld
    clean_overlayed_slide osld
    clean_empty_slide osld
    
    Dim shapes_arraylist As Object
    Dim shapes_no_lines As Object
    Dim shapes_font_size As Object
    Dim shapes_names As Object
    
    Dim oshp As Variant
    
    Dim total_fixed_height As Double
    Dim available_vertical As Double
    Dim sum As Double
    Dim leftThird As Integer
    Dim rightThird As Integer
    Dim ternary_search_left As Integer
    Dim title_shape As Variant
    Dim title_line_Spacing As Double
    Dim target_percentage As Double
    Dim initial_bottom As Double
    Dim padding_amount As Double
    
    lines_ub = get_lines_ub()
    lines_lb = get_lines_lb()
    precision = get_line_spacing_precision()
    
    Set shapes_arraylist = CreateObject("System.Collections.ArrayList")

    'loop on each shape

    For Each oshp In get_ungroupedshapes(osld.shapes)
            Set extended = New extended_shape
            Set extended.shape = oshp
            Set extended.shape_bbox = get_shape_bbox(oshp)
            extended.shape_name = oshp.Name
            If Not is_fixed(extended) Then
                shapes_arraylist.Add extended
            End If
    Next oshp
        
    Set groups_arraylist = get_sorted_shapes_groups(shapes_arraylist, osld.shapes)
    reformat_text_shprng_hidden osld.shapes, depth, groups_arraylist
 
 
    shapes_arraylist.Clear
    If get_slide_type(osld) = "title slide" Then
        Set title_shape = osld.shapes.Item(1)
        title_shape.TextFrame.MarginBottom = 3.6
        title_shape.TextFrame.MarginLeft = 7.2
        title_shape.TextFrame.MarginRight = 7.2
        title_shape.TextFrame.MarginTop = 3.6
        title_shape.TextFrame.AutoSize = ppAutoSizeShapeToFitText
        title_shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
        title_shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
              
        title_shape.Width = 0.78 * osld.CustomLayout.Width
        
        title_shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
        title_shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 2
        
        title_shape.Height = 1.5 * heights(get_i_of_line_spacing(2), get_textbox_number_of_lines(heights, title_shape), get_i_of_font(title_shape.TextFrame.TextRange.Characters(0, 1).font.Size))
        title_shape.Left = (osld.CustomLayout.Width - title_shape.Width) / 2
        title_shape.Top = (osld.CustomLayout.Height - title_shape.Height) / 2
    Else
        For Each oshp_extended In groups_arraylist
           Set oshp = oshp_extended.shape
           Set extended = New extended_shape
           Set extended.shape = oshp
           If get_shape_type(oshp) = "textbox" Then
              oshp.TextFrame.MarginBottom = 3.6
              oshp.TextFrame.MarginLeft = 7.2
              oshp.TextFrame.MarginRight = 7.2
              oshp.TextFrame.MarginTop = 3.6
              oshp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
              oshp.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
              oshp.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
             
              oshp.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
              oshp.TextFrame.TextRange.ParagraphFormat.SpaceWithin = get_line_spacing_i(1) ' initial line spacing is small as possible to get rid of overlapping textboxes
                If TypeName(title_shape) = "Empty" Then  'is_title_text_box(oshp) And
                       Set title_shape = oshp
                Else
                     oshp.Left = 0.03 * osld.CustomLayout.Width + (0.94 * osld.CustomLayout.Width - oshp.Width) ' leave only 3% on the left :D
                     
                     
                     
                     before = get_textbox_number_of_lines(heights, oshp)
                     nearest = nearest_width(oshp)
                     oshp.ScaleWidth nearest / oshp.Width, msoFalse, msoScaleFromBottomRight
                     oshp.TextFrame.WordWrap = True
                     after = get_textbox_number_of_lines(heights, oshp)
                     If before < after Then
                         nearest = get_next_width(oshp)
                         oshp.ScaleWidth nearest / oshp.Width, msoFalse, msoScaleFromBottomRight
                         oshp.TextFrame.WordWrap = True
                     End If
                     
                End If
                
                
                Set extended.shape_bbox = get_shape_bbox(oshp)
                extended.shape_name = oshp.Name
                shapes_arraylist.Add extended
           Else
                Set extended.shape_bbox = get_shape_bbox(oshp)
                extended.shape_name = oshp.Name
                If Not is_fixed(extended) Then
                    shapes_arraylist.Add extended
                End If
           End If
            
        Next oshp_extended
        

        
        If shapes_arraylist.Count > 1 Then
            
            get_title_width title_shape
            title_shape.Left = (osld.CustomLayout.Width - title_shape.Width) / 2
            
             
            sort_shapes shapes_arraylist
            'shapes are now sorted
            
            Set shapes_no_lines = CreateObject("System.Collections.ArrayList")
            Set shapes_font_size = CreateObject("System.Collections.ArrayList")
            
            target_percentage = 0
            total_fixed_height = 0
            For Each ex_oshp In shapes_arraylist
                If Not is_fixed(ex_oshp) Then ' filter shapes
                    Set oshp = ex_oshp.shape
                    If get_shape_type(oshp) = "textbox" Then
                        shapes_font_size.Add (get_i_of_font(oshp.TextFrame.TextRange.Characters(0, 1).font.Size))
                        shape_no_lines = get_textbox_number_of_lines(heights, oshp)
                        shapes_no_lines.Add (shape_no_lines)
                        target_percentage = target_percentage + shape_no_lines
                    Else
                        total_fixed_height = total_fixed_height + oshp.Height
                    End If
                End If
            Next ex_oshp
            
            target_percentage = max(min(target_percentage * 0.2 + shapes_arraylist.Count() * 2, 15), 15)

            
            
            top_position = get_initial_top_position(osld)
            title_shape.Top = top_position
   
            sum = get_hieght_sum(heights, get_i_of_line_spacing(2.5), shapes_no_lines, shapes_font_size, total_fixed_height)
            
            initial_bottom = get_initial_bottom_position(osld)
            
            bottom_position = min(initial_bottom, top_position + sum / (1 - target_percentage / 100)) '0.963 = 1 - 0.037 which is from the bottom
            
            'round up bottom position
            If (initial_bottom - bottom_position) < initial_bottom * 0.35 And initial_bottom <> bottom_position Then
                title_shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
                title_shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3
                title_shape.Top = initial_top_position
                title_hieght = title_shape.Height
                 
                avail_old = bottom_position - top_position
                avail_new = initial_bottom - top_position
                padding_old = avail_old * target_percentage / 100
                padding_new = (initial_bottom - bottom_position) / 3 + padding_old
                target_percentage = padding_new / avail_new * 100
                bottom_position = initial_bottom
            ElseIf (initial_bottom - bottom_position) > initial_bottom * 0.35 Then
                title_shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
                title_shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3
                target_percentage = target_percentage * 1.2
                bottom_position = bottom_position * 1.54
            End If
            
            available_vertical = bottom_position - top_position
            target_percentage = max(8, 1.3 * (available_vertical - total_fixed_height) * target_percentage / available_vertical)
            
            If shapes_arraylist.Count - 1 < 3 Then
                target_percentage = target_percentage / (4 - shapes_arraylist.Count) / 1.5
            End If
            
            ternary_search_left = 1
            ternary_search_right = get_line_spacing_precision()
            While ternary_search_left < ternary_search_right
                third = Int((ternary_search_right - ternary_search_left) / 3)
                leftThird = ternary_search_left + third
                rightThird = ternary_search_right - third
                
                cost_leftThird = convex_cost_function_for_alignment(target_percentage, leftThird, heights, available_vertical, shapes_no_lines, shapes_font_size, total_fixed_height)
                cost_rightThird = convex_cost_function_for_alignment(target_percentage, rightThird, heights, available_vertical, shapes_no_lines, shapes_font_size, total_fixed_height)
                
                If cost_leftThird < cost_rightThird Then
                    If ternary_search_right = rightThird Then
                        ternary_search_right = rightThird - 1
                    Else
                        ternary_search_right = rightThird
                    End If
                    
                Else
                    If ternary_search_left = leftThird Then
                        ternary_search_left = leftThird + 1
                    Else
                        ternary_search_left = leftThird
                    End If
                    
                End If
            Wend
            ' ternary_search_left and ternary_search_right are equal
            
           
            
            padding_percentage = get_padding_percentage(ternary_search_left, heights, available_vertical, shapes_no_lines, shapes_font_size, total_fixed_height)
            padding_amount = (padding_percentage / 100 * available_vertical) / shapes_arraylist.Count
            
            If shapes_arraylist.Count - 1 > 4 Then
                padding_amount_new = min(0.5 * heights(get_i_of_line_spacing(2.2), 1, shapes_font_size.Item(0)), padding_amount)
                delta_padding_amount = padding_amount - padding_amount_new
                bottom_position = bottom_position - delta_padding_amount * shapes_arraylist.Count
                padding_amount = padding_amount_new
            End If
            

            shape_index = 0
            textbox_index = 0
            Set shapes_names = CreateObject("System.Collections.ArrayList")
            For Each ex_oshp In shapes_arraylist
        
                If get_shape_type(ex_oshp.shape) = "textbox" Then
                    ex_oshp.shape.LockAspectRatio = False
                    If shapes_no_lines.Item(textbox_index) = 1 And get_line_spacing_i(ternary_search_left) > 2.5 Then
                        ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
                        ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 2.2
                        ex_oshp.shape.Height = heights(get_i_of_line_spacing(2.2), shapes_no_lines.Item(textbox_index), shapes_font_size.Item(textbox_index))
                    Else
                        ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = msoTrue
                        ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = get_line_spacing_i(ternary_search_left)
                        ex_oshp.shape.Height = heights(ternary_search_left, shapes_no_lines.Item(textbox_index), shapes_font_size.Item(textbox_index))
                    End If
                    If textbox_index = 0 Then
                        ex_oshp.shape.Top = top_position
                    End If
                    
                    textbox_index = textbox_index + 1
                
                Else
                   
                    ex_oshp.shape.Left = (osld.CustomLayout.Width - ex_oshp.shape.Width) / 2
                        
                End If
                If shapes_arraylist.Count - 1 = shape_index Then ' last shape
                  ex_oshp.shape.Top = bottom_position - ex_oshp.shape.Height
                ElseIf shape_index = 1 Then ' the shape after the title shape
                    ex_oshp.shape.Top = get_title_space_multiplier() * padding_amount + top_position + title_shape.Height
                ElseIf shapes_arraylist.Count - 1 > shape_index And shape_index > 1 Then
                  ex_oshp.shape.Top = shapes_arraylist.Item(shape_index - 1).shape.Top + shapes_arraylist.Item(shape_index - 1).shape.Height
                End If
        
                shape_index = shape_index + 1
                shapes_names.Add (ex_oshp.shape.Name)
            Next ex_oshp
            
            'size corretion
            shape_index = 1
            For Each oshp In shapes_arraylist
                If get_shape_type(oshp.shape) = "textbox" Then
                    oshp.shape.Height = oshp.shape.Height + (padding_percentage - target_percentage) * available_vertical / (100 * (shapes_arraylist.Count))
                    
                    If shape_index = shapes_arraylist.Count And oshp.shape.Height + oshp.shape.Top > initial_bottom Then
                        oshp.shape.Top = oshp.shape.Top - (padding_percentage - target_percentage) * available_vertical / (100 * (shapes_arraylist.Count))
                    End If
                        
                        ' for things hard to fit in slide
                    If shapes_arraylist.Count = 1 Then
                        remaining_for_object = initial_bottom - (title_shape.Top + title_shape.Height) - oshp.shape.Height
                        If oshp.shape.Height + oshp.shape.Top > initial_bottom Then
                            If remaining_for_object > 0 Then
                                oshp.shape.Top = remaining_for_object / 2 + title_shape.Top + title_shape.Height
                            Else
                                oshp.shape.Top = title_shape.Top + title_shape.Height
                            End If
                        End If
                        
                        If get_shape_type(oshp.shape) = "table" And remaining_for_object > 0 Then
                            oshp.shape.Top = remaining_for_object / 2 + title_shape.Top + title_shape.Height
                        End If
                    
                    End If
                End If
                
                  
                shape_index = shape_index + 1
            Next oshp
            
            ' fixing wide shapes
            shape_index = 0
            textbox_index = 0
            shift_up = 0
            For Each ex_oshp In shapes_arraylist
                ex_oshp.shape.Top = ex_oshp.shape.Top - shift_up
                
                If get_shape_type(ex_oshp.shape) = "textbox" Then
                    If shapes_no_lines.Item(textbox_index) = 1 And ex_oshp.shape.Height > 1.15 * heights(get_i_of_line_spacing(2.2), 1, shapes_font_size.Item(textbox_index)) Then
                        shift_up = shift_up + ex_oshp.shape.Height - 1.15 * heights(get_i_of_line_spacing(2.2), shapes_no_lines.Item(textbox_index), shapes_font_size.Item(textbox_index))
                        ex_oshp.shape.Height = 1.15 * heights(get_i_of_line_spacing(2.2), shapes_no_lines.Item(textbox_index), shapes_font_size.Item(textbox_index))
                    End If
                    
                    If shapes_no_lines.Item(textbox_index) > 1 And shape_index = 0 And ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin > 1.75 Then
                        h1 = ex_oshp.shape.Height
                        
                        ex_oshp.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.5
                        ex_oshp.shape.Top = top_position
                        
                        h2 = ex_oshp.shape.Height
                        shift_up = shift_up + h1 - h2
                    End If
                    textbox_index = textbox_index + 1
                End If
            
                shape_index = shape_index + 1
            Next ex_oshp
           
            
            If delta_padding_amount = 0 Then
                shapes_names.removeat (0)
            End If
            
            If shapes_names.Count > 2 Then
                 osld.shapes.Range(shapes_names.toArray()).Distribute msoDistributeVertically, False
            End If
            
     
    
            ' ungroup everything
            get_ungroupedshapes osld.shapes
        End If
    End If
End Function




