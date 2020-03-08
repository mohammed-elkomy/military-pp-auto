Attribute VB_Name = "toolbar_textrange"
'here a build toolbar used for text based operations
'coloring, fonts, setting last visible character to zero size

Const toolbarname As String = " ‰”Ìﬁ «·ŒÿÊÿ"

Sub text_toolbar()
    Dim tooltips(1 To 9) As String
    For i = 1 To UBound(get_text_colors())
        tooltips(i) = " œ—Ã «·Œÿ " & i
    Next i
  
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton

    'on error go to
    On Error GoTo escape
    CommandBars(toolbarname).Delete
escape:
    
    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)

    For i = 1 To UBound(get_text_colors())
        ' Now add a button to the new toolbar
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
        ' And set some of the button's properties
    
        With oButton
             .DescriptionText = tooltips(i)
              'Tooltip text when mouse if placed over button
              
             .Caption = tooltips(i)
             'Text if Text in Icon is chosen
    
             .OnAction = "text_color_" & i
             
    
             .Style = msoButtonIcon
              ' Button displays as icon, not text or both

             .Picture = LoadPicture(get_root_dir & "toolbars\f" & i & ".jpg")
        End With
    
    Next i
    
      
    ' add reduce font button
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oButton
         .DescriptionText = " ’€Ì— «·Œÿ"
         .Caption = " ’€Ì— «·Œÿ"
         .OnAction = "reduce_font"
         .Style = msoButtonIcon
          .FaceId = 63
    End With
    
    ' add increase font button
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oButton
         .DescriptionText = " ﬂ»Ì— «·Œÿ"
         .Caption = " ﬂ»Ì— «·Œÿ"
         .OnAction = "increase_font"
         .Style = msoButtonIcon
          .FaceId = 62
    End With
    

    
    ' add text font button
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oButton
         .DescriptionText = "‰Ê⁄ «·Œÿ"
         .Caption = "‰Ê⁄ «·Œÿ"
         .OnAction = "text_font" ' pt bold heading
         .Style = msoButtonIcon
          .FaceId = 3989
    End With
     
    oToolbar.Visible = True
    inner_colors_disable
End Sub


Sub text_color_1()
   On Error GoTo showerror
    'shaperange_correction
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(1)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(1)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(1)
                        End If
                    Next Cell
                Next Row
                
            End If
            
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_2()
    On Error GoTo showerror
    'shaperange_correction
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(2)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(2)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(2)
                        End If
                    Next Cell
                Next Row
                
            End If
            
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_3()
   On Error GoTo showerror
    'shaperange_correction

    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(3)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(3)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(3)
                        End If
                    Next Cell
                Next Row
                
            End If
            
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_4()
   On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(4)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(4)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(4)
                        End If
                    Next Cell
                Next Row
            End If
            
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_5()
    On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(5)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(5)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(5)
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_6()
On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(6)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(6)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(6)
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_7()
On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(7)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(7)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(7)
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_8()
On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(8)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(8)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(8)
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub text_color_9()
On Error GoTo showerror
    
    text_colors = get_text_colors()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        ActiveWindow.Selection.TextRange.font.color = text_colors(9)
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                shape.TextFrame.TextRange.font.color = text_colors(9)
            ElseIf get_shape_type(shape) = "table" Then
                For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                            Cell.shape.TextFrame.TextRange.font.color = text_colors(9)
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub
Sub text_font()
    On Error GoTo showerror
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        With ActiveWindow.Selection.TextRange.font
            .NameAscii = "PT Bold Heading"
            .NameComplexScript = "PT Bold Heading"
            .Bold = msoFalse
        End With
        ActiveWindow.Selection.TextRange.ParagraphFormat.Alignment = 7 ' justify low
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                With shape.TextFrame.TextRange.font
                    .NameAscii = "PT Bold Heading"
                    .NameComplexScript = "PT Bold Heading"
                    .Bold = msoFalse
                End With
                
                 shape.TextFrame.TextRange.ParagraphFormat.Alignment = 7 ' justify low
            ElseIf get_shape_type(shape) = "table" Then
                 For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                             With Cell.shape.TextFrame.TextRange.font
                                .NameAscii = "PT Bold Heading"
                                .NameComplexScript = "PT Bold Heading"
                                .Bold = msoFalse
                             End With
                              Cell.shape.TextFrame.TextRange.ParagraphFormat.Alignment = 7 ' justify low
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub reduce_font()
On Error GoTo showerror
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        With ActiveWindow.Selection.TextRange.font
            .Size = .Size - 1
        End With
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                With shape.TextFrame.TextRange.Characters(1, get_utf_string_length(shape.TextFrame.TextRange.text) - 1).font
                    .Size = .Size - 1
                End With
            ElseIf get_shape_type(shape) = "table" Then
                 For Each Row In shape.Table.rows
                    For Each Cell In Row.Cells
                        If Cell.Selected Then
                             With Cell.shape.TextFrame.TextRange.Characters(1, get_utf_string_length(Cell.shape.TextFrame.TextRange.text)).font
                                .Size = .Size - 1
                             End With
                        End If
                    Next Cell
                Next Row
            End If
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub

Sub increase_font()
    On Error GoTo showerror
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        With ActiveWindow.Selection.TextRange.font
            .Size = .Size + 1
        End With
    Else:
        For Each shape In ActiveWindow.Selection.shaperange
            If get_shape_type(shape) = "textbox" Then
                   With shape.TextFrame.TextRange.Characters(1, get_utf_string_length(shape.TextFrame.TextRange.text) - 1).font
                       .Size = .Size + 1
                   End With
            ElseIf get_shape_type(shape) = "table" Then
                   For Each Row In shape.Table.rows
                        For Each Cell In Row.Cells
                           If Cell.Selected Then
                                With Cell.shape.TextFrame.TextRange.Characters(1, get_utf_string_length(Cell.shape.TextFrame.TextRange.text)).font
                                    .Size = .Size + 1
                                End With
                              
                           End If
                        Next Cell
                   Next Row
            End If
        Next shape
    End If
    GoTo fin
showerror:
        show_error "ÕœÀ Œÿ√ Ê  „ ≈Ã—«¡ „Õ«Ê·… ·≈’·«ÕÂ"
        fix_office_bug ActiveWindow.View.slide
fin:
End Sub



Sub text_colors_enable()
    enable_bar toolbarname
End Sub

Sub text_colors_disable()
    disble_bar toolbarname
End Sub





