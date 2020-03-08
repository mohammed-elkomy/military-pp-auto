Attribute VB_Name = "toolbar_inner_colors"
Const toolbarname As String = "«·Ê«‰ «·»‰Êœ"


Sub inner_colors_toolbar()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    
    Dim tooltips(0 To 10) As String
    For i = 0 To UBound(get_text_colors())
        tooltips(i) = " œ—Ã «·»‰œ " & i
    Next i
    
    'on error go to
    On Error GoTo escape
    CommandBars(toolbarname).Delete
escape:
    
    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, Position:=msoBarRight, Temporary:=False)
    
    For i = 0 To min(10, UBound(get_dark_gradient_colors()))
        ' Now add a button to the new toolbar
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
        ' And set some of the button's properties
        With oButton
    
             .DescriptionText = tooltips(i)
              'Tooltip text when mouse if placed over button
    
             .Caption = tooltips(i)
             'Text if Text in Icon is chosen
    
             .OnAction = "inner_dark_color_gradient_" & i
             
    
             .Style = msoButtonIcon
              ' Button displays as icon, not text or both

              .Picture = LoadPicture(get_root_dir & "toolbars\dark-gradient\" & (i + 1) & ".jpg")
        End With
    Next i
    
    For i = 0 To min(10, UBound(get_light_gradient_colors()))
        ' Now add a button to the new toolbar
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
        ' And set some of the button's properties
        With oButton
    
             .DescriptionText = tooltips(i)
              'Tooltip text when mouse if placed over button
    
             .Caption = tooltips(i)
             'Text if Text in Icon is chosen
    
             .OnAction = "inner_light_color_gradient_" & i
             
    
             .Style = msoButtonIcon
              ' Button displays as icon, not text or both

              .Picture = LoadPicture(get_root_dir & "toolbars\light-gradient\" & (i + 1) & ".jpg")
        End With
    Next i
    
    For i = 0 To min(10, UBound(get_inner_dark_colors()))
        ' Now add a button to the new toolbar
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
        ' And set some of the button's properties
    
        With oButton
    
             .DescriptionText = tooltips(i)
              'Tooltip text when mouse if placed over button
    
             .Caption = tooltips(i)
             'Text if Text in Icon is chosen
    
             .OnAction = "inner_color_dark_" & i
             
    
             .Style = msoButtonIcon
              ' Button displays as icon, not text or both

              .Picture = LoadPicture(get_root_dir & "toolbars\dark\" & (i + 1) & ".jpg")
        End With
    
    Next i
    
    For i = 0 To min(10, UBound(get_inner_light_colors()))
        ' Now add a button to the new toolbar
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
        ' And set some of the button's properties
    
        With oButton
    
             .DescriptionText = tooltips(i)
              'Tooltip text when mouse if placed over button

             .Caption = tooltips(i)
             'Text if Text in Icon is chosen
    
             .OnAction = "light_inner_color_" & i
             
    
             .Style = msoButtonIcon
              ' Button displays as icon, not text or both

              .Picture = LoadPicture(get_root_dir & "toolbars\light\" & (i + 1) & ".jpg")
        End With
    
    Next i
    oToolbar.Visible = True
    inner_colors_disable
End Sub

Function outline_color()
    For Each oshp In get_linear_shapes(ActiveWindow.Selection.shaperange)
            If get_shape_type(oshp) = "picture" Or get_shape_type(oshp) = "other" Or get_shape_type(oshp) = "group" Then
                oshp.Line.ForeColor.rgb = 255
            End If
    Next oshp
End Function


Sub inner_light_color_gradient(color As Variant)
    On Error GoTo escape:
     border_color = get_light_gardient_border_rgb()
     
    'shaperange_correction
    Set tables = from_non_title_to_title()
   'fix_office_bug ActiveWindow.View.slide.shapes, ActiveWindow.View.slide
    
    For Each Table In tables
        For Each Row In Table.Table.rows
            For Each Cell In Row.Cells
                If Cell.Selected Then
                    
                     With Cell.shape
                         .Fill.Transparency = 0#
                         .Fill.TwoColorGradient msoGradientHorizontal, 3
                     
                     End With
                     
                     With Cell.shape
                        .Fill.ForeColor.rgb = rgb_div(color)
                        .Fill.BackColor.rgb = color
                        .Fill.TwoColorGradient msoGradientHorizontal, 3
                     End With
                End If
            Next Cell
        Next Row
    Next Table
   

        
         With ActiveWindow.Selection.shaperange
             .Line.ForeColor.SchemeColor = ppBackground
             .Line.Visible = msoTrue
             .Fill.Transparency = 0#
             .Fill.TwoColorGradient msoGradientHorizontal, 3
             .Line.ForeColor.rgb = border_color
             .Line.Visible = msoTrue
             .Line.Weight = 3#
         End With
         
         With ActiveWindow.Selection.shaperange
            .Fill.ForeColor.rgb = rgb_div(color)
            .Fill.BackColor.rgb = color
            .Fill.TwoColorGradient msoGradientHorizontal, 3
            .Line.ForeColor.rgb = border_color
            .Line.Visible = msoTrue
            .Line.Weight = 3#
        End With
 
    
    
escape:
    If is_correct_shape(ActiveWindow.Selection.shaperange.Item(1)) Then
        outline_color
    Else:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
    End If
End Sub

Sub inner_light_color_gradient_0()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(0)
End Sub

Sub inner_light_color_gradient_1()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(1)
End Sub

Sub inner_light_color_gradient_2()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(2)
End Sub

Sub inner_light_color_gradient_3()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(3)
End Sub

Sub inner_light_color_gradient_4()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(4)
End Sub

Sub inner_light_color_gradient_5()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(5)
End Sub

Sub inner_light_color_gradient_6()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(6)
End Sub

Sub inner_light_color_gradient_7()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(7)
End Sub

Sub inner_light_color_gradient_8()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(7)
End Sub

Sub inner_light_color_gradient_9()
    colours = get_light_gradient_colors()
    inner_light_color_gradient colours(7)
End Sub

Sub inner_dark_color_gradient(color As Variant)
    On Error GoTo escape:
    'shaperange_correction
    border_color = get_dark_gardient_border_rgb()
    
    Set shapes_names = CreateObject("System.Collections.ArrayList")
    Set tables = CreateObject("System.Collections.ArrayList")
    
    For Each oshp In get_linear_shapes(ActiveWindow.Selection.shaperange)
            If get_shape_type(oshp) = "textbox" Then
                shapes_names.Add oshp.Name
                
                If get_shape_color_type(oshp) = "light" Then ' light background to dark background
                    For i = 1 To oshp.TextFrame.TextRange.length
                        If oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) Then
                            oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                        Else
                            oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 0) ' make other things yellow
                        End If
                    Next i
                End If
            ElseIf get_shape_type(oshp) = "table" Then
                 tables.Add oshp
                 
                 For Each Row In oshp.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                              If get_shape_color_type(Cell.shape) = "light" Then ' light background to dark background
                                For i = 1 To Cell.shape.TextFrame.TextRange.length
                                    If Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) Then
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                                    Else
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 0) ' make other things yellow
                                    End If
                                Next i
                            End If
                        End If
                    Next Cell
                Next Row
            End If
    Next oshp
    
    For Each Table In tables
        For Each Row In Table.Table.rows
            For Each Cell In Row.Cells
                If Cell.Selected Then
                    
                     With Cell.shape
                         .Fill.Transparency = 0#
                         .Fill.TwoColorGradient msoGradientHorizontal, 3
                     
                     End With
                     
                     With Cell.shape
                        .Fill.ForeColor.rgb = rgb_div(color)
                        .Fill.BackColor.rgb = color
                        .Fill.TwoColorGradient msoGradientHorizontal, 3
                     End With
                End If
            Next Cell
        Next Row
   Next Table
   
    If shapes_names.Count > 0 Then
         ActiveWindow.View.slide.shapes.Range(shapes_names.toArray()).Select
         With ActiveWindow.Selection.shaperange
             .Line.ForeColor.SchemeColor = ppBackground
             .Line.Visible = msoTrue
             .Fill.Transparency = 0#
             .Fill.TwoColorGradient msoGradientHorizontal, 3
             .Line.ForeColor.rgb = border_color
             .Line.Visible = msoTrue
             .Line.Weight = 3#
         End With
         
         With ActiveWindow.Selection.shaperange
            .Fill.ForeColor.rgb = rgb_div(color)
            .Fill.BackColor.rgb = color
            .Fill.TwoColorGradient msoGradientHorizontal, 3
            .Line.ForeColor.rgb = border_color
            .Line.Visible = msoTrue
            .Line.Weight = 3#
        End With
    End If
    

escape:
    If is_correct_shape(ActiveWindow.Selection.shaperange.Item(1)) Then
        outline_color
    Else:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
    End If
End Sub


Sub inner_dark_color_gradient_0()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(0)
End Sub

Sub inner_dark_color_gradient_1()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(1)
End Sub

Sub inner_dark_color_gradient_2()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(2)
End Sub

Sub inner_dark_color_gradient_3()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(3)
End Sub

Sub inner_dark_color_gradient_4()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(4)
End Sub

Sub inner_dark_color_gradient_5()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(5)
End Sub

Sub inner_dark_color_gradient_6()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(6)
End Sub

Sub inner_dark_color_gradient_7()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(7)
End Sub

Sub inner_dark_color_gradient_8()
    colours = get_dark_gradient_colors()
    inner_dark_color_gradient colours(7)
End Sub

Sub inner_dark_color_gradient_9()
    colours = get_dark_gradient_colors()
    inner_color_dark_gradient colours(7)
End Sub

Sub inner_color_dark(color As Variant)
    On Error GoTo escape:
    'shaperange_correction
    border_color = get_dark_border_rgb()
     
    Set shapes_names = CreateObject("System.Collections.ArrayList")
    Set tables = CreateObject("System.Collections.ArrayList")
    
    For Each oshp In get_linear_shapes(ActiveWindow.Selection.shaperange)
            If get_shape_type(oshp) = "textbox" Then
                shapes_names.Add oshp.Name
                
                 If get_shape_color_type(oshp) = "light" Then ' light background to dark background
                    For i = 1 To oshp.TextFrame.TextRange.length
                        If oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) Then
                            oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                        Else
                            oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 0) ' make other things yellow
                        End If
                    Next i
                End If
            ElseIf get_shape_type(oshp) = "table" Then
                 tables.Add oshp
                 
                 For Each Row In oshp.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                             If get_shape_color_type(Cell.shape) = "light" Then ' light background to dark background
                                For i = 1 To Cell.shape.TextFrame.TextRange.length
                                    If Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) Then
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                                    Else
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 0) ' make other things yellow
                                    End If
                                Next i
                            End If
                        End If
                    Next Cell
                Next Row
            End If
    Next oshp
    
    For Each Table In tables
        For Each Row In Table.Table.rows
            For Each Cell In Row.Cells
                If Cell.Selected Then
                     With Cell.shape
                        .Fill.ForeColor.SchemeColor = ppFill
                        .Fill.ForeColor.rgb = color
                        .Fill.Solid
                     End With
                     
                     With Cell.shape
                        .Fill.ForeColor.SchemeColor = ppFill
                        .Fill.ForeColor.rgb = color
                        .Fill.Solid
                    End With
                End If
            Next Cell
        Next Row
   Next Table
   
    
    With ActiveWindow.Selection.shaperange
        .Line.ForeColor.rgb = border_color
        .Fill.ForeColor.SchemeColor = ppFill
        .Fill.ForeColor.rgb = color
        .Line.Visible = msoTrue
        .Fill.Solid
    End With
    
    With ActiveWindow.Selection.shaperange
        .Line.ForeColor.rgb = border_color
        .Fill.ForeColor.SchemeColor = ppFill
        .Fill.ForeColor.rgb = color
        .Line.Visible = msoTrue
        .Fill.Solid
    End With
escape:
    If is_correct_shape(ActiveWindow.Selection.shaperange.Item(1)) Then
        outline_color
    Else:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
    End If
End Sub

Sub inner_color_dark_0()
    colours = get_inner_dark_colors()
    inner_color_dark colours(0)
End Sub

Sub inner_color_dark_1()
    colours = get_inner_dark_colors()
    inner_color_dark colours(1)
End Sub

Sub inner_color_dark_2()
    colours = get_inner_dark_colors()
    inner_color_dark colours(2)
End Sub

Sub inner_color_dark_3()
    colours = get_inner_dark_colors()
    inner_color_dark colours(3)
End Sub

Sub inner_color_dark_4()
    colours = get_inner_dark_colors()
    inner_color_dark colours(4)
End Sub

Sub inner_color_dark_5()
    colours = get_inner_dark_colors()
    inner_color_dark colours(5)
End Sub

Sub inner_color_dark_6()
    colours = get_inner_dark_colors()
    inner_color_dark colours(6)
End Sub

Sub inner_color_dark_7()
    colours = get_inner_dark_colors()
    inner_color_dark colours(7)
End Sub

Sub inner_color_dark_8()
    colours = get_inner_dark_colors()
    inner_color_dark colours(8)
End Sub

Sub inner_color_dark_9()
    colours = get_inner_dark_colors()
    inner_color_dark colours(9)
End Sub


Function from_non_title_to_title()
    Dim tables As Variant
    Set shapes_names = CreateObject("System.Collections.ArrayList")
    Set tables = CreateObject("System.Collections.ArrayList")
     
    For Each oshp In get_linear_shapes(ActiveWindow.Selection.shaperange)
        If get_shape_type(oshp) = "textbox" Then
            shapes_names.Add oshp.Name
                
            If get_shape_color_type(oshp) = "dark" Then ' dark background to light background
                For i = 1 To oshp.TextFrame.TextRange.length
                    If oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) Then
                        oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0)  ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                    Else
                        oshp.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 0, 0) ' make other things yellow
                    End If
                Next i
            End If
        ElseIf get_shape_type(oshp) = "table" Then
            For Each Row In oshp.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            If get_shape_color_type(Cell.shape) = "dark" Then ' dark background to light background
                                For i = 1 To Cell.shape.TextFrame.TextRange.length
                                    If Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 255, 255) Then
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the oshp.textframe.textrange.Characters(i, 1).font.color.rgb = rgb(0, 0, 0) ' make the black text  white
                                    Else
                                        Cell.shape.TextFrame.TextRange.Characters(i, 1).font.color.rgb = rgb(255, 0, 0) ' make other things yellow
                                    End If
                                Next i
                            End If
                        End If
                    Next Cell
                Next Row
                
            tables.Add oshp
        End If
    Next oshp
    
    If shapes_names.Count > 0 Then
        ActiveWindow.View.slide.shapes.Range(shapes_names.toArray()).Select
    End If
    
    Set from_non_title_to_title = tables
End Function


Sub light_inner_color_0()
    colours = get_inner_light_colors()
    light_inner_color colours(0)
End Sub

Sub light_inner_color_1()
    colours = get_inner_light_colors()
    light_inner_color colours(1)
End Sub
Sub light_inner_color_2()
    colours = get_inner_light_colors()
    light_inner_color colours(2)
End Sub

Sub light_inner_color_3()
    colours = get_inner_light_colors()
    light_inner_color colours(3)
End Sub

Sub light_inner_color_4()
    colours = get_inner_light_colors()
    light_inner_color colours(4)
End Sub

Sub light_inner_color_5()
    colours = get_inner_light_colors()
    light_inner_color colours(5)
End Sub

Sub light_inner_color_6()
    colours = get_inner_light_colors()
    light_inner_color colours(6)
End Sub


Sub light_inner_color_7()
    colours = get_inner_light_colors()
    light_inner_color colours(6)
End Sub

Sub light_inner_color_8()
    colours = get_inner_light_colors()
    light_inner_color colours(6)
End Sub

Sub light_inner_color_9()
    colours = get_inner_light_colors()
    light_inner_color colours(6)
End Sub

Sub light_inner_color(color As Variant)
    On Error GoTo escape:
     border_color = get_light_border_rgb()
     
    'shaperange_correction
    Set tables = from_non_title_to_title()
    
    For Each Table In tables
        For Each Row In Table.Table.rows
            For Each Cell In Row.Cells
                If Cell.Selected Then
                     With Cell.shape
                        .Fill.ForeColor.SchemeColor = ppFill
                        .Fill.ForeColor.rgb = color
                        .Fill.Solid
                     End With
                     
                     With Cell.shape
                        .Fill.ForeColor.SchemeColor = ppFill
                        .Fill.ForeColor.rgb = color
                        .Fill.Solid
                    End With
                End If
            Next Cell
        Next Row
   Next Table
   'fix_office_bug ActiveWindow.View.slide.shapes, ActiveWindow.View.slide
    
    With ActiveWindow.Selection.shaperange
        .Line.ForeColor.rgb = border_color
        .Fill.ForeColor.SchemeColor = ppFill
        .Fill.ForeColor.rgb = color
        .Fill.Visible = msoTrue
        .Fill.Solid
    End With
    
     With ActiveWindow.Selection.shaperange
        .Line.ForeColor.rgb = border_color
        .Fill.ForeColor.SchemeColor = ppFill
        .Fill.ForeColor.rgb = color
        .Fill.Visible = msoTrue
        .Fill.Solid
    End With
escape:
    If is_correct_shape(ActiveWindow.Selection.shaperange.Item(1)) Then
        outline_color
    Else:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
    End If
End Sub

Sub inner_colors_enable()
    enable_bar toolbarname
End Sub

Sub inner_colors_disable()
    disble_bar toolbarname
End Sub


