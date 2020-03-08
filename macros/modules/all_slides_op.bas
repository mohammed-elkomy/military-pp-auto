Attribute VB_Name = "all_slides_op"
Function word_highlighter(pattern As String, color As Long)
    ' coloring numbers and dates based on regex
    '''''''''''''''''''''''''''''''''''''''''
    Set regX = CreateObject("vbscript.regexp")
    With regX
    .Global = True
    .pattern = pattern
    End With
    
    ' loop on each slide
    For Each osld In ActivePresentation.Slides
    'loop on each shape
        For Each oshp In get_linear_shapes(osld.shapes)  '   osld.shapes
            ' get the text range if available
             If get_shape_type(oshp) = "textbox" Then
                Set oTxtRng = oshp.TextFrame.TextRange
                
                ' find and replace
                Set myMatches = regX.Execute(oTxtRng.text)
                
                For Each mymatch In myMatches
                    ' works for office 2010
                    Set otmprng = oTxtRng.Characters(InStr(oTxtRng.text, mymatch.Value), mymatch.length)
       
                    ' a check for office 2003 bug :3
                    If otmprng.text <> mymatch.Value Then
                        Set otmprng = oTxtRng.Characters(mymatch.firstindex, mymatch.length)
                    End If
                    
                    If get_shape_color_type(oshp) = "dark" Then
                        If oshp.Fill.Visible Then
                        otmprng.font.color.rgb = rgb(255, 255, 0) ' yellow for white-text titles (when background colors are left as is)
                        Else
                        otmprng.font.color.rgb = color   ' red for black-text titles (when background color is removed)
                        End If
                    Else
                        otmprng.font.color.rgb = color ' red for numbers in normal text boxes
                    End If
                Next
            End If
        Next oshp
    Next osld
End Function



Sub removing_interior_foreground_color()
    ' used for inkjet 2800 printer
    '''''''''''''''''''''''''''''''''''''''''
    make_progressor
      
    For Each slide In ActivePresentation.Slides
        With slide
            .FollowMasterBackground = msoFalse
            .DisplayMasterShapes = msoTrue
            With .Background
                .Fill.Visible = msoTrue
                .Fill.ForeColor.SchemeColor = ppBackground
                .Fill.Transparency = 0#
                .Fill.Solid
            End With
        End With
    Next slide
    
    ' loop on each slide
    slide_count = ActivePresentation.Slides.Count
    For Each osld In ActivePresentation.Slides
        'loop on each shape in osld
        For Each oshp In get_linear_shapes(osld.shapes)
            ' get the text range if available
            If get_shape_type(oshp) = "textbox" Then
                    If oshp.Fill.ForeColor.rgb <> rgb(255, 255, 255) And oshp.Fill.Visible Then ' if arleady have interior color
                        If get_shape_color_type(oshp) = "dark" Then ' for things in main titles
                            oshp.Line.ForeColor.rgb = rgb(255, 0, 0) ' make the borders red
                            
                            For i = 1 To oshp.TextFrame.TextRange.length
                                'If oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) Then
                                '    oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the white text black
                                'End If
                                
                                oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) - oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb
                                
                                With oshp.TextFrame.TextRange.Characters(i, 1).font.color
                                    If .rgb = rgb(255, 255, 0) Then
                                         .rgb = rgb(200, 200, 0) ' yellow to gold
                                    End If
                                    
                                    If .rgb = rgb(255, 255, 255) Then
                                        .rgb = rgb(0, 0, 0) ' white to black
                                    End If
                                End With
                            Next i
                        Else
                            oshp.Line.ForeColor.rgb = oshp.Fill.ForeColor.rgb ' copy outer color
                            
                            If oshp.Line.ForeColor.rgb = rgb(255, 255, 0) Then
                                oshp.Line.ForeColor.rgb = rgb(200, 200, 0)
                            End If
                            
                            If oshp.Line.ForeColor.rgb = rgb(255, 255, 255) Then
                                oshp.Line.ForeColor.rgb = rgb(0, 0, 0) ' white to black
                            End If
                            
                            
                            
                        End If
                        
                        oshp.Fill.Solid
                        oshp.Fill.ForeColor.rgb = rgb(255, 255, 255)
                    End If
                
            ElseIf get_shape_type(oshp) = "table" Then
                    
                    For Each Row In oshp.Table.rows
                            For Each Cell In Row.Cells
                                 With Cell.shape
                                 
                                    '.Fill.Transparency = 1
                                    If .Fill.ForeColor.rgb <> rgb(255, 255, 255) Then
                                        '>>>> copying outer color
                                        '.TextFrame.TextRange.font.color = max(255, .Fill.ForeColor.rgb)
                                        
                                        'If .TextFrame.TextRange.font.color = rgb(255, 255, 0) Then
                                        '    .TextFrame.TextRange.font.color = rgb(200, 200, 0)
                                        'End If
                                        '>>> controling colors
                                        
                                        
                                      If (.Fill.GradientVariant <> 0 Or .Fill.ForeColor.rgb = rgb(255, 0, 0)) Then
                                        For i = 1 To .TextFrame.TextRange.length
                                           
                                           .TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) - .TextFrame.TextRange.Characters(i, 1).font.color.rgb
                                        
                                        Next i
                                      End If
                                      
                                       For i = 1 To .TextFrame.TextRange.length
                                           With .TextFrame.TextRange.Characters(i, 1).font.color
                                               If .rgb = rgb(255, 255, 0) Then
                                                   .rgb = rgb(200, 200, 0) ' yellow to gold
                                               End If
                                               
                                               If .rgb = rgb(255, 255, 255) Then
                                                   .rgb = rgb(0, 0, 0) ' white to black
                                               End If
                                           End With
                                        Next i
                                        
                                    End If
                                End With
                                    
                            Next Cell
                    Next Row
                    
                    oshp.Fill.Solid
                    oshp.Fill.ForeColor.rgb = rgb(255, 255, 255)
            End If
            
            If oshp.Type = msoAutoShape Then
                solid_bw oshp
                fill_bw oshp
            End If
        
        Next oshp
        squeeze_height_shprng osld.shapes
        
        update_progressor Int(osld.SlideIndex / slide_count * 100)
        DoEvents
    Next osld
    
      
    done_progressor
End Sub



Sub line_bw(oshp As Variant)
    On Error GoTo fin
    oshp.Line.ForeColor.rgb = rgb(0, 0, 0) ' make the borders black
fin:
End Sub

Sub solid_bw(oshp As Variant)
    On Error GoTo fin
    oshp.Fill.Solid
fin:
End Sub

Sub fill_bw(oshp As Variant)
    On Error GoTo fin
    oshp.Fill.ForeColor.rgb = rgb(255, 255, 255)
fin:
End Sub

Sub removing_interior_foreground_color_black_white()
    ' used for inkjet 2800 printer
    '''''''''''''''''''''''''''''''''''''''''
    make_progressor
      
    For Each slide In ActivePresentation.Slides
        With slide
            .FollowMasterBackground = msoFalse
            .DisplayMasterShapes = msoTrue
            With .Background
                .Fill.Visible = msoTrue
                .Fill.ForeColor.SchemeColor = ppBackground
                .Fill.Transparency = 0#
                .Fill.Solid
            End With
        End With
    Next slide
    
    ' loop on each slide
    slide_count = ActivePresentation.Slides.Count
    For Each osld In ActivePresentation.Slides
        'loop on each shape in osld
        For Each oshp In get_linear_shapes(osld.shapes)
            ' get the text range if available
            If get_shape_type(oshp) = "textbox" Then
                    oshp.Line.ForeColor.rgb = rgb(0, 0, 0) ' make the borders black
                    oshp.TextFrame.TextRange.font.color.rgb = rgb(0, 0, 0)  ' make text black
                    
                    oshp.Fill.Solid
                    oshp.Fill.ForeColor.rgb = rgb(255, 255, 255)
                    
    
            ElseIf get_shape_type(oshp) = "table" Then
                    For Each Row In oshp.Table.rows
                            For Each Cell In Row.Cells
                                 With Cell.shape
                                 
                                    '.Fill.Transparency = 1
                                    .TextFrame.TextRange.font.color = rgb(0, 0, 0)
                                 End With
                                 For Each border In Cell.Borders
                                     border.ForeColor.rgb = rgb(0, 0, 0)
                                 Next border
                            Next Cell
                    Next Row
                    oshp.Fill.Solid
                    oshp.Fill.ForeColor.rgb = rgb(255, 255, 255)
                    
                    'oshp.Line.ForeColor.SchemeColor = ppForeground
                    'oshp.Line.Visible = msoTrue
            End If
            If oshp.Type = msoAutoShape Then
                line_bw oshp
                solid_bw oshp
                fill_bw oshp
            End If
            
        Next oshp
        squeeze_height_shprng osld.shapes
        update_progressor Int(osld.SlideIndex / slide_count * 100)
        DoEvents
    Next osld

    done_progressor
End Sub

Sub primary_animation(s_start As Integer, s_end As Integer)
    
    make_progressor
    DoEvents
    ' this will find and sort(based on top location considering orientation and groups things in one line for grouped animation) all textboxes within a slide then adds blind entry animation to them
    '-------------------------------------------------------------
    ' loop on each slide
    Dim current_queue As Variant
    Dim next_queue As Variant
    Dim last_bbox As vertical_projection_bbox
    Set last_bbox = New vertical_projection_bbox
    
    Set current_queue = CreateObject("System.Collections.Queue")
    Set next_queue = CreateObject("System.Collections.Queue")
    Set slide_range_numbers = CreateObject("System.Collections.ArrayList")
    
    For i = s_start To s_end
        slide_range_numbers.Add (i)
    Next i
    
    slide_count = s_end - s_start + 1
    
    SlideIndex = 0
    
    For Each osld In ActivePresentation.Slides.Range(slide_range_numbers.toArray())
        Debug.Print "starting" & osld.SlideIndex
        
        last_delay = 0
        SlideIndex = SlideIndex + 1
        ' box effect
         With osld.SlideShowTransition
            .AdvanceOnClick = msoTrue
            .AdvanceOnTime = msoFalse
            .AdvanceTime = 0
            .Duration = 0.5
            .EntryEffect = ppEffectBoxOut
            .Hidden = msoFalse
            .Speed = ppTransitionSpeedFast
        End With
        

        Set shapes_arraylist = CreateObject("System.Collections.ArrayList")
        'loop on each shape

        For Each oshp In get_ungroupedshapes(osld.shapes)
           Set extended = New extended_shape
           Set extended.shape = oshp
           Set extended.shape_bbox = get_shape_bbox(oshp)
           extended.shape_name = oshp.Name
           shapes_arraylist.Add extended
           DoEvents
        Next oshp
        
        Set groups_arraylist = get_sorted_shapes_groups(shapes_arraylist, osld.shapes)
  
        shape_index = 1
        animate_index = 1
        match = True
        For Each ex_oshp In groups_arraylist
            Set oshp = ex_oshp.shape
            
            If get_shape_type(oshp) = "textbox" Then
                next_queue.enqueue oshp.TextFrame.TextRange.text
            End If
            
            If current_queue.Count > 0 And get_shape_type(oshp) = "textbox" And match Then
                match = (oshp.TextFrame.TextRange.text = current_queue.dequeue())
            Else
                match = False
            End If
            DoEvents
            
            If get_shape_type(oshp) = "textbox" Or get_shape_type(oshp) = "picture" Or get_shape_type(oshp) = "table" Or get_shape_type(oshp) = "other" Or get_shape_type(oshp) = "group" Then
                If shapes_arraylist.Count() = 1 Or match Or shape_index = 1 Then
                    ' title text box and repeated textboxes
                    With oshp.AnimationSettings
                        .AdvanceMode = 1
                        .AfterEffect = 0
                        .Animate = False ' disable animation
                        .AnimateBackground = 0
                        .AnimateTextInReverse = 0
                        .EntryEffect = 257
                        .TextLevelEffect = 0
                        .TextUnitEffect = 0
                    End With
                Else
                    
                    Set shapes_to_sort = CreateObject("System.Collections.ArrayList")
                    
                    Set linears = get_sub_ungroupedshapes(osld.shapes, oshp)
                    
                    'loop on each shape
                    For Each oshp_in_grp In linears
                       Set extended = New extended_shape
                       Set extended.shape = oshp_in_grp
                       Set extended.shape_bbox = get_shape_vertical_projection(oshp_in_grp)
                       extended.shape_name = oshp_in_grp.Name
                       shapes_to_sort.Add extended
                       DoEvents
                    Next oshp_in_grp
        
                    sort_shapes shapes_to_sort
                    
                    For Each single_shape In shapes_to_sort
                         ' non title text box
                        With single_shape.shape.AnimationSettings
                            .Animate = True ' enable animation
                            .EntryEffect = 769 ' blinds
                            .AdvanceMode = 2 ' after previous
                            .AfterEffect = 0
                            .AnimateBackground = -1
                            .AnimateTextInReverse = 0
                            .AnimationOrder = animate_index '(one-based)
                            .TextLevelEffect = 16
                            .TextUnitEffect = 0
                        End With
                        Application.StartNewUndoEntry
                        If animate_index > 1 Then
                            last_bbox.max_y = min(last_bbox.max_y, single_shape.shape_bbox.max_y)
                            delay_intersection_complement = 1 - single_shape.shape_bbox.vertical_jaccard(last_bbox)
                            If delay_intersection_complement > 1 Then
                                delay_intersection_complement = 1
                            End If
                        Else
                            delay_intersection_complement = 0
                        End If
                        
                       
                        osld.TimeLine.MainSequence.Item(animate_index).Timing.TriggerDelayTime = last_delay + delay_intersection_complement * 0.5

                        last_delay = last_delay + delay_intersection_complement * 0.5
                        Set last_bbox = single_shape.shape_bbox
                        animate_index = animate_index + 1 ' for the next object
                        DoEvents
                    Next single_shape
                End If
            End If
            DoEvents
            shape_index = shape_index + 1 ' for the next object
        Next ex_oshp
        
        Set current_queue = next_queue
        Set next_queue = CreateObject("System.Collections.Queue")
        
        For i = osld.TimeLine.MainSequence.Count To 1 Step -1
            osld.TimeLine.MainSequence.Item(i).Timing.TriggerType = msoAnimTriggerWithPrevious
        Next i
        
        
        
        update_progressor Int(SlideIndex / slide_count * 100)
        DoEvents
    Next osld
    done_progressor
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub primary_animation_archived()
    ' this will find and sort(based on top location considering orientation and groups things in one line for grouped animation) all textboxes within a slide then adds blind entry animation to them
    '-------------------------------------------------------------
    ' loop on each slide
    Dim current_queue As Variant
    Dim next_queue As Variant
    Set current_queue = CreateObject("System.Collections.Queue")
    Set next_queue = CreateObject("System.Collections.Queue")
    For Each osld In ActivePresentation.Slides
        ' box effect
        
         With osld.SlideShowTransition
            .AdvanceOnClick = msoTrue
            .AdvanceOnTime = msoFalse
            .AdvanceTime = 0
            .Duration = 0.5
            .EntryEffect = ppEffectBoxOut
            .Hidden = msoFalse
            .Speed = ppTransitionSpeedFast
        End With
        

        Set shapes_arraylist = CreateObject("System.Collections.ArrayList")
        'loop on each shape

        For Each oshp In get_ungroupedshapes(osld.shapes)
           Set extended = New extended_shape
           Set extended.shape = oshp
           Set extended.shape_bbox = get_shape_bbox(oshp)
           extended.shape_name = oshp.Name
           shapes_arraylist.Add extended
        Next oshp
        
        sort_shapes shapes_arraylist
 
        For i = 0 To sorted_array.Count() - 1:
            Set sorted_array.Item(i).shape_bbox = get_shape_vertical_projection(sorted_array.Item(i).shape)
            If sublist.Count() > 0 Then
                If group_bbox.vertical_jaccard(sorted_array.Item(i).shape_bbox) < 0.6 And sorted_array.Item(i).shape_bbox.vertical_jaccard(group_bbox) < 0.6 Then
                    ret_list.Add sublist
                    Set sublist = CreateObject("System.Collections.ArrayList")
                End If
            Else
                group_bbox.min_y = sorted_array.Item(i).shape_bbox.min_y
                group_bbox.max_y = sorted_array.Item(i).shape_bbox.max_y
            End If
            sublist.Add sorted_array.Item(i)
            
            group_bbox.min_y = min(sorted_array.Item(i).shape_bbox.min_y, group_bbox.min_y)
            group_bbox.max_y = max(sorted_array.Item(i).shape_bbox.max_y, group_bbox.max_y)
        Next i
    

        shape_index = 1
        match = True
        For Each ex_oshp In groups_arraylist
            Set oshp = ex_oshp.shape
            
            If get_shape_type(oshp) = "textbox" Then
                next_queue.enqueue oshp.TextFrame.TextRange.text
            End If
            
            If current_queue.Count > 0 And get_shape_type(oshp) = "textbox" And match Then
                match = (oshp.TextFrame.TextRange.text = current_queue.dequeue())
            Else
                match = False
            End If
            
            If get_shape_type(oshp) = "textbox" Or get_shape_type(oshp) = "picture" Or get_shape_type(oshp) = "table" Or get_shape_type(oshp) = "other" Or get_shape_type(oshp) = "group" Then
                If shapes_arraylist.Count() = 1 Or match Then
                    ' title text box
                    With oshp.AnimationSettings
                        .AdvanceMode = 1
                        .AfterEffect = 0
                        .Animate = False ' disable animation
                        .AnimateBackground = 0
                        .AnimateTextInReverse = 0
                        .EntryEffect = 257
                        .TextLevelEffect = 0
                        .TextUnitEffect = 0
                    End With
                Else
                    ' non title text box
                    With oshp.AnimationSettings
                        .Animate = True ' enable animation
                        .EntryEffect = 769 ' blinds
                        .AdvanceMode = 2 ' after previous
                        .AfterEffect = 0
                        .AnimateBackground = -1
                        .AnimateTextInReverse = 0
                        .AnimationOrder = shape_index '(one-based)
                        .TextLevelEffect = 16
                        .TextUnitEffect = 0
                    End With
                    
                    shape_index = shape_index + 1 ' for the next object
                End If
            End If
            
        Next ex_oshp
        
        Set current_queue = next_queue
        Set next_queue = CreateObject("System.Collections.Queue")
        
    Next osld
    
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub secondary_animation(s_start As Integer, s_end As Integer)
    make_progressor
    slide_count = s_end - s_start + 1
    
    Set slide_range_numbers = CreateObject("System.Collections.ArrayList")
    
    For i = s_start To s_end
        slide_range_numbers.Add (i)
    Next i
    
    
    ' loop on each slide
    For Each osld In ActivePresentation.Slides.Range(slide_range_numbers.toArray())
         With osld.SlideShowTransition
            .AdvanceOnClick = msoTrue
            .AdvanceOnTime = msoFalse
            .AdvanceTime = 0
            .Duration = 0.5
            .EntryEffect = ppEffectBoxOut
            .Hidden = msoFalse
            .Speed = ppTransitionSpeedFast
        End With
        
        'loop on each shape in osld
        For Each oshp In osld.shapes
            ' get the text range if available
            oshp.AnimationSettings.Animate = False ' disable animation
        Next oshp
        update_progressor Int(osld.SlideIndex / slide_count * 100)
        DoEvents
    Next osld
    done_progressor
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub clean_empty_slide(osld As Variant)
    'loop on each shape in osld
    For Each this_shp In get_linear_shapes(osld.shapes)
        ' get the text range if available
        shape_type = get_shape_type(this_shp)
        If shape_type = "empty" Or (shape_type = "textbox" And this_shp.Height * this_shp.Width / (osld.CustomLayout.Height * osld.CustomLayout.Width) < 0.001) Then
            this_shp.Cut
        End If
        
        'If shape_type = "other" Then
        '    this_shp.ZOrder msoSendToBack
        'End If
        DoEvents
    Next this_shp
End Sub

Sub clean_empty(offset)
    slide_count = ActivePresentation.Slides.Count
    For Each osld In ActivePresentation.Slides
      
        clean_empty_slide osld
        update_progressor Int(osld.SlideIndex / slide_count * 33 + offset)
        DoEvents
    Next osld
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub clean_overlayed_slide(osld As Variant)
    'loop on each shape in osld
    For Each this_shp In get_linear_shapes(osld.shapes)
        ' get the text range if available
        If get_shape_type(this_shp) = "textbox" Then
                this_shp.Shadow.Visible = False
                
                Dim this_bbox As bounding_box
                Set this_bbox = get_shape_bbox(this_shp)
                
                For Each that_shp In get_linear_shapes(osld.shapes)
                     If that_shp.Id <> this_shp.Id Then
                         Dim that_bbox As bounding_box
                         Set that_bbox = get_shape_bbox(that_shp)
                         If this_bbox.jaccard(that_bbox) > 0.9 And that_shp.ZOrderPosition < this_shp.ZOrderPosition Then
                            that_shp.Cut
                         End If
                     End If
                Next that_shp
        End If
    Next this_shp
End Sub

Sub clean_overlayed(offset)
    slide_count = ActivePresentation.Slides.Count
    ' loop on each slide
    For Each osld In ActivePresentation.Slides
        clean_overlayed_slide osld
        update_progressor Int(osld.SlideIndex / slide_count * 33 + offset)
        DoEvents
    Next osld
End Sub

Sub adjust_corners_slide(osld As Variant)
     'loop on each shape in osld
     For Each this_shp In get_linear_shapes(osld.shapes)
         ' get the text range if available
    
         If get_shape_type(this_shp) = "textbox" Then
             this_shp.AutoShapeType = msoShapeRoundedRectangle
             If this_shp.Adjustments.Count > 0 Then
                 this_shp.Adjustments(1) = 0.07
             End If
             
             
             With this_shp.TextFrame
                 .HorizontalAnchor = msoAnchorNone
                 .VerticalAnchor = msoAnchorMiddle
             End With
                
         End If
        DoEvents
     Next this_shp
        
End Sub

Sub adjust_corners(offset)
    slide_count = ActivePresentation.Slides.Count
    ' loop on each slide
    For Each osld In ActivePresentation.Slides
        adjust_corners_slide osld
        update_progressor Int(osld.SlideIndex / slide_count * 33 + offset)
        DoEvents
    Next osld
End Sub

Function cleanrs()
    make_progressor
    
    adjust_corners (0)
    clean_overlayed (33)
    clean_empty (66)
    
    done_progressor
End Function

