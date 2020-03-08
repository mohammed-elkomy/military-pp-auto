Attribute VB_Name = "shape_selection_op"
Function replace_special_char(oTxtRng As Variant, regX As Variant, placeholder As String)
    '' find and replace
    Set myMatches = regX.Execute(oTxtRng)
    delta = 0
    
    If myMatches.Count < 50 Then
        For Each mymatch In myMatches
            ' works for office 2010
            Set otmprng = oTxtRng.Characters(mymatch.firstindex + 1 + delta, mymatch.length)
            length_before = otmprng.length
            otmprng.text = " " & placeholder & " "
            delta = delta + 3 - length_before
        Next
    End If
    
     
End Function

 Function replace_special_chars(oshp As Variant, regexes As Variant, special_chars As Variant)
        Dim my_char As String
        regex_index = 0
        For Each regex In regexes
            my_char = Mid(special_chars(regex_index), Len(special_chars(regex_index)), 1)
            replace_special_char oshp.TextFrame.TextRange, regex, my_char
            regex_index = regex_index + 1
        Next regex

 End Function
 
Function reformat_text_shprng(shapes_range As Variant)
    reformat_text_shprng_hidden shapes_range, 0, "nothing"
End Function

Function fix_bad_ascii(oshp As Variant, actual_length As Integer)
    
        
    For i = 1 To actual_length
       Select Case Mid(oshp.TextFrame.TextRange, i, 1)
          Case ChrW(8211)
              oshp.TextFrame.TextRange.Characters(start:=i, length:=1).text = "-"
          'Case Chr(13)
              'oshp.textframe.textrange.Characters(Start:=i, length:=1).text = Chr(11)
              'Dim sub_str As String
              
              'sub_str = oshp.TextFrame.textrange.Characters(Start:=i, length:=2).text
              'leng = 2
              'For j = 1 To get_utf_string_length(sub_str)
              '     If Asc(Mid(sub_str, j, 1)) > 33 Then
              '         leng = 1
              '     End If
              'Next j
              
              'oshp.TextFrame.textrange.Characters(Start:=i, length:=2).text = Chr(11) + "Ú"
       End Select
    Next i
End Function


Function reformat_text_shprng_hidden(shapes_range As Variant, depth As Integer, groups_arraylist As Variant)
    ' depth is either 0,1 or 2 an increase in depth means decrease in fonts
    ' the default fonts are set with depth = 0
 
    Dim my_char As String
    Dim special_chars() As Variant
    Dim actual_length As Integer
    
    Set regexes = CreateObject("System.Collections.ArrayList")
    special_chars = Array("/", "-") ' , "\[", "\]",  "\(", "\)"
    
    For Each special_char In special_chars
        'regexes.Add (make_regex("\s*" & special_char & "\s*"))
        regexes.Add (make_regex("[\u0020]*" & special_char & "[\u0020]*"))
    Next special_char
    
    slide_type = get_slide_type(shapes_range.Parent)
    
    'loop on each shape
    For Each oshp In get_linear_shapes(shapes_range)
        If get_shape_type(oshp) = "textbox" Then
            oshp.TextFrame.TextRange.ParagraphFormat.Alignment = 7 ' justify low
             
            If depth >= 0 Then
                If slide_type = "title slide" Then
                    oshp.TextFrame.TextRange.Characters(1, get_utf_string_length(oshp.TextFrame.TextRange.text) - 1).font.Size = get_next_smaller_font(get_main_title_font_size(), depth)
                Else
                    If TypeName(groups_arraylist) = "ArrayList" Then
                        If groups_arraylist.Item(0).shape.Name = oshp.Name Then
                            oshp.TextFrame.TextRange.Characters(1, get_utf_string_length(oshp.TextFrame.TextRange.text) - 1).font.Size = get_next_smaller_font(get_header_font_size(), depth)
                        Else
                            oshp.TextFrame.TextRange.Characters(1, get_utf_string_length(oshp.TextFrame.TextRange.text) - 1).font.Size = get_next_smaller_font(get_inner_font_size(), depth)
                        End If
                    End If
                    
                 
                End If
            End If
            
            actual_length = get_utf_string_length(oshp.TextFrame.TextRange.text)
            replace_special_chars oshp, regexes, special_chars
            fix_bad_ascii oshp, actual_length

            last_index = -1
            For i = min_int(actual_length, Len(oshp.TextFrame.TextRange)) To 1 Step -1
                If Asc(Mid(oshp.TextFrame.TextRange, i, 1)) > 33 And last_index = -1 Then ' finding the last character index
                    last_index = i + 1
                End If
            Next i
        
            If last_index <= actual_length Then
                 remaining = actual_length - last_index + 1 ' trimming white space chars
                 oshp.TextFrame.TextRange.Characters(last_index, remaining) = Chr(11)
                 oshp.TextFrame.TextRange.Characters(last_index, 1).font.Size = 1
            Else ' we need to add some characters

                 oshp.TextFrame.TextRange.Characters(actual_length, 1).text = oshp.TextFrame.TextRange.Characters(actual_length, 1).text & Chr(11)
                 oshp.TextFrame.TextRange.Characters(last_index, 1).font.Size = 1
            End If
        ElseIf get_shape_type(oshp) = "table" Then
            For Each Row In oshp.Table.rows
                For Each Cell In Row.Cells
                    If Cell.Selected Then
                        actual_length = get_utf_string_length(Cell.shape.TextFrame.TextRange.text)
                        replace_special_chars Cell.shape, regexes, special_chars
                        fix_bad_ascii Cell.shape, actual_length
                    End If
                Next Cell
            Next Row
                
        End If
    Next oshp
End Function


Function squeeze_height_shprng(shapes_range As Variant)
    'loop on each shape
    For Each oshp In get_linear_shapes(shapes_range)
        If get_shape_type(oshp) = "textbox" Then
                   
            last_index = -1
            actual_length = get_utf_string_length(oshp.TextFrame.TextRange.text)
             
                For i = actual_length To 1 Step -1
                    If Asc(Mid(oshp.TextFrame.TextRange, i, 1)) > 33 And last_index = -1 Then ' finding the last character index
                        last_index = i + 1
                    End If
                Next i
            
                If last_index <= actual_length Then
                     remaining = actual_length - last_index + 1 ' trimming white space chars
                     oshp.TextFrame.TextRange.Characters(last_index, remaining) = Chr(11)
                     oshp.TextFrame.TextRange.Characters(last_index, 1).font.Size = 1
                Else ' we need to add some characters
    
                     oshp.TextFrame.TextRange.Characters(actual_length, 1).text = oshp.TextFrame.TextRange.Characters(actual_length, 1).text & Chr(11)
                     oshp.TextFrame.TextRange.Characters(last_index, 1).font.Size = 1
                End If
 
        End If
    Next oshp
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function nearest_innercolor_shprng(shapes_range As Variant, colours() As Long)
    'loop on each shape in osld
    For Each oshp In get_linear_shapes(shapes_range)
        
        ' get the text range if available
        If get_shape_type(oshp) = "textbox" Then
            
                Dim min_cost As Long
                min_cost = 256
                min_cost = min_cost * min_cost * min_cost * 100
                min_index = -1
                If oshp.Fill.Visible Then ' if arleady have interior color
                    If oshp.Fill.GradientVariant = 0 Then
                    
                    
                        For i = 1 To UBound(colours)
                            dist = rgb_distance(colours(i), oshp.Fill.ForeColor.rgb)
                            Debug.Print dist
                            If dist < min_cost Then
                                min_cost = dist
                                min_index = i
                            End If
                        Next i
                        
                        oshp.Fill.ForeColor.rgb = colours(min_index)
                    End If
                End If
        End If
    Next oshp
End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function parenthesized_text_colorization_shape(oshp As Variant, colours As Variant, header_colours As Variant)
    ' create a stack for parenthesis
    Dim parentheses_stack As Object
    Set parentheses_stack = CreateObject("System.Collections.Stack")
    
    ' create a stack for nested executions
    Set exec_stack = CreateObject("System.Collections.Stack")
    
    For counter = 1 To Len(oshp.TextFrame.TextRange.text)
        current_char = Mid(oshp.TextFrame.TextRange.text, counter, 1)
        
        If current_char = "{" Or _
        current_char = "}" Or _
        current_char = "(" Or _
        current_char = ")" Or _
        current_char = "]" Or _
        current_char = "[" Then
            ' check for braces/braket/parenthesis
            Dim current_char_pp As TextRange
            Set current_char_pp = oshp.TextFrame.TextRange.Characters(counter, 1)
        
            If Not is_arabic(current_char_pp) Then
                Select Case current_char
                    Case "{"
                        current_char = "}"
                    Case "}"
                        current_char = "{"
                    Case "("
                        current_char = ")"
                    Case ")"
                        current_char = "("
                    Case "]"
                        current_char = "["
                    Case "["
                        current_char = "]"
                End Select
            End If
            
            If parentheses_stack.Count > 0 Then ' if stack not empty
                stack_top = Asc(parentheses_stack.PEEK().parenthesis)
            Else
                stack_top = -1000
            End If
            
            Dim ostack_entry As New stack_entry
            If stack_top + Asc(current_char) = Asc(")") + Asc("(") Or _
            stack_top + Asc(current_char) = Asc("[") + Asc("]") Or _
            stack_top + Asc(current_char) = Asc("}") + Asc("{") Then
                ' check for opening and closing if not empty
                perenthesized_length = counter - parentheses_stack.PEEK().startIndex - 1
                
                For current_start_index = parentheses_stack.PEEK().startIndex + 1 To 1 Step -1 ' search space [max(-inf,startindex),+1]
                    If oshp.TextFrame.TextRange.Characters(current_start_index, perenthesized_length).text = Mid(oshp.TextFrame.TextRange.text, parentheses_stack.PEEK().startIndex + 1, perenthesized_length) Then
                        Dim oexec_entry As New exec_entry
                        Set oexec_entry = New exec_entry
                        oexec_entry.depth = parentheses_stack.Count
                        Set oexec_entry.sub_textrange = oshp.TextFrame.TextRange.Characters(current_start_index, perenthesized_length)
                        exec_stack.push oexec_entry
                        Exit For
                    End If
                Next
                parentheses_stack.pop
            Else
                ' push new entry if no balance achieved
                Set ostack_entry = New stack_entry
                ostack_entry.startIndex = counter ' starting index
                ostack_entry.parenthesis = current_char ' add the braces/braket/parenthesis
                parentheses_stack.push ostack_entry ' push that pair to stack
            End If
        End If
    Next
    
    
    While exec_stack.Count > 0
        Set oexec_entry = exec_stack.pop()
        If get_shape_color_type(oshp) = "dark" Then 'kokokomy
            oexec_entry.sub_textrange.font.color.rgb = header_colours(oexec_entry.depth + 2)
        Else
            oexec_entry.sub_textrange.font.color.rgb = colours(oexec_entry.depth + 2)
        End If
        
        
    Wend
End Function

Function parenthesized_text_colorization_shprng(shapes_range As Variant)
    Dim colours() As Long
    Dim header_colours() As Long
    
    colours = get_text_colors()
    header_colours = get_header_text_colors()
    
    'loop on each shape
    For Each oshp In get_linear_shapes(shapes_range)
        
        ' get the text range if available
        If get_shape_type(oshp) = "textbox" Then
            parenthesized_text_colorization_shape oshp, colours, header_colours
            
        ElseIf get_shape_type(oshp) = "table" Then
            For Each Row In oshp.Table.rows
                For Each Cell In Row.Cells
                    If Cell.Selected Then
                        parenthesized_text_colorization_shape Cell.shape, colours, header_colours
                    End If
                Next Cell
            Next Row
        End If
    Next oshp
    
    'coloring_numbers_and_dates_shprng shapes_range, get_numbers_and_dates_regex
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function coloring_numbers_and_dates_shape(oshp As Variant, regX As Variant)
    current_start = 1
    Set oTxtRng = oshp.TextFrame.TextRange
                
    ' find and replace
    Set myMatches = regX.Execute(oTxtRng.text)
    
    For Each mymatch In myMatches
        ' works for office 2010
        current_start = current_start + InStr(Mid(oTxtRng.text, current_start), mymatch.Value) - 1
        Set otmprng = oTxtRng.Characters(current_start, mymatch.length)
        current_start = current_start + mymatch.length
        
        If get_shape_color_type(oshp) = "dark" Then
            If oshp.Fill.Visible Then
                otmprng.font.color.rgb = rgb(255, 255, 0) ' yellow for white-text titles (when background colors are left as is)
            Else
                otmprng.font.color.rgb = rgb(255, 0, 0)   ' red for black-text titles (when background color is removed)
            End If
        Else
            otmprng.font.color.rgb = rgb(255, 0, 0) ' red for numbers in normal text boxes
        End If
        
    Next
End Function

Function coloring_numbers_and_dates_shprng(shapes_range As Variant, regX As Variant)
    
    reformat_text_shprng_hidden shapes_range, -1, "nothing"
    'loop on each shape
    For Each oshp In get_linear_shapes(shapes_range)
        ' get the text range if available
        If get_shape_type(oshp) = "textbox" Then
            coloring_numbers_and_dates_shape oshp, regX
        ElseIf get_shape_type(oshp) = "table" Then
            For Each Row In oshp.Table.rows
                For Each Cell In Row.Cells
                    If Cell.Selected Then
                        coloring_numbers_and_dates_shape Cell.shape, regX
                    End If
                Next Cell
            Next Row
        End If
    Next oshp
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


