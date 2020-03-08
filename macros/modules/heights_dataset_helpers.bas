Attribute VB_Name = "heights_dataset_helpers"
Function big_array_generator()
    ' 400(spacing),30(lines),10(font) one-based-indexing tensor
    Const line_spacing_precision As Integer = 300
    Const maximum_number_of_lines As Integer = 25
    Const number_of_fonts As Integer = 32 - 10 + 1
    
    Debug.Assert get_line_spacing_precision() = line_spacing_precision
    Debug.Assert get_maximum_number_of_lines() = maximum_number_of_lines
    Debug.Assert get_number_of_fonts() = number_of_fonts

    Dim heights(1 To line_spacing_precision, 1 To maximum_number_of_lines, 1 To number_of_fonts) As Double
    big_array_generator = heights
End Function

Function generate_heights()
    heights = big_array_generator()
    Set shape = ActiveWindow.View.slide.shapes.Item(1)
    
    With shape.TextFrame
        .MarginBottom = 3.6
        .MarginLeft = 7.2
        .MarginRight = 7.2
        .MarginTop = 3.6
        .AutoSize = ppAutoSizeShapeToFitText
        .TextRange.ParagraphFormat.SpaceAfter = 0
        .TextRange.ParagraphFormat.SpaceBefore = 0
        .TextRange.text = ""
    End With
    
    For k = 1 To get_number_of_fonts()
        shape.TextFrame.TextRange.text = ""
        For j = 1 To get_maximum_number_of_lines()
            shape.TextFrame.TextRange.text = shape.TextFrame.TextRange.text & "«»Ã ÂÊ“" & Str(j) & Chr(11)
            shape.TextFrame.TextRange.Characters(0, Len(shape.TextFrame.TextRange.text) - 1).font.Size = get_font_size_i(k)
            shape.TextFrame.TextRange.Characters(Len(shape.TextFrame.TextRange.text), 1).font.Size = 1
            
            For i = 1 To get_line_spacing_precision()
               shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = get_line_spacing_i(i)
               heights(i, j, k) = shape.Height
            Next i
         Next j
    Next k
    generate_heights = heights
End Function


Function generate_heights_file()
    'disable screen updating
    ScreenUpdating = False
    ' --- Long time consuming code

    heights = generate_heights()
    file_path = Application.ActivePresentation.Path & "\heights.komy"
    Open file_path For Output As #1
    For i = 1 To get_line_spacing_precision()
        For j = 1 To get_maximum_number_of_lines()
            For k = 1 To get_number_of_fonts()
                'Print #1, "heights(" & i & "," & j & "," & k & ")=" & heights(i, j, k) & Chr(13)
                Print #1, i & "," & j & "," & k & "," & heights(i, j, k)
            Next k
        Next j
    Next i
    Close #1
    ' Redraw screen again
    ScreenUpdating = True
  
End Function

Function read_heights_file()
    
    heights = big_array_generator()
    file_path = get_root_dir & "heights.komy"
    Open file_path For Input As #2
    Dim buffer As String
    While Not EOF(2)
        Line Input #2, buffer
        pieces = Split(buffer, ",")
        i = CInt(pieces(0))
        j = CInt(pieces(1))
        k = CInt(pieces(2))
        heights(i, j, k) = CDbl(pieces(3))
        'Debug.Assert Abs(heights(i, j, k) - CDbl(pieces(3)))) < 0.001
    Wend
    Close #2
    read_heights_file = heights
End Function


Function show_min_spacing()
    heights = read_heights_file()
    
    Dim min_spacing As Double
    min_spacing = 255
    min_spacing = min_spacing * min_spacing
    For i = 1 To get_line_spacing_precision()
        For j = 1 To get_maximum_number_of_lines() - 1
            For k = 1 To get_number_of_fonts()
                min_spacing = min(heights(i, j + 1, k) - heights(i, j, k), min_spacing)
            Next k
        Next j
    Next i
    
    Debug.Print min_spacing
    
End Function


Function find_threshold(line_spacing As Integer, font As Integer, heights() As Double)
    Dim min_spacing As Double
    min_spacing = 255
    min_spacing = min_spacing * min_spacing
    
    For j = 1 To get_maximum_number_of_lines() - 1
        min_spacing = min(heights(line_spacing, j + 1, font) - heights(line_spacing, j, font), min_spacing)
    Next j
    find_threshold = min_spacing / 2
End Function



Function get_line_spacing_i(ByVal i As Integer) As Double
    'expected value in range 1 ~ get_line_spacing_precision
    lines_ub = get_lines_ub()
    lines_lb = get_lines_lb()
    precision = get_line_spacing_precision()
    get_line_spacing_i = lines_lb + (i - 1) * ((lines_ub - lines_lb) / (precision - 1))
End Function

Function get_i_of_line_spacing(ByVal line_spacing As Double) As Integer
    'expected value in range .6 ~ 3.2
    lines_ub = get_lines_ub()
    lines_lb = get_lines_lb()
    precision = get_line_spacing_precision()
    Step = ((lines_ub - lines_lb) / (precision - 1))
    
    get_i_of_line_spacing = Int((line_spacing - lines_lb) / Step) + 1

End Function

Function get_font_size_i(ByVal i As Integer) As Integer
    ' available font sizes 10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32
    get_font_size_i = i + 9
End Function

Function get_i_of_font(ByVal font As Integer) As Integer
    ' available font sizes 10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32
    get_i_of_font = font - 9
End Function

Function get_next_smaller_font(fontsize As Integer, depth As Integer)
    get_next_smaller_font = get_font_size_i(get_i_of_font(fontsize) - depth)
End Function

Function get_textbox_number_of_lines2(heights() As Double, shape As Variant) 'binary finder,deprecated
    Dim line_spacing As Integer
    Dim font_size As Integer
    Dim threshold As Double
    
    search_key = shape.Height
    line_spacing = get_i_of_line_spacing(shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin)
    font_size = get_i_of_font(shape.TextFrame.TextRange.font.Size)
    threshold = find_threshold(line_spacing, font_size, heights)
    
    binary_search_lower = 1
    binary_search_upper = get_maximum_number_of_lines()
    
    While binary_search_lower <= binary_search_upper
        binary_search_mid = Int((binary_search_upper + binary_search_lower) / 2)
        
        If Abs(heights(line_spacing, binary_search_mid, font_size) - search_key) < threshold Then
            get_textbox_number_of_lines = binary_search_mid
            
            Exit Function
        ElseIf heights(line_spacing, binary_search_mid, font_size) < search_key Then
            binary_search_lower = binary_search_mid + 1
        Else
            binary_search_upper = binary_search_mid - 1
        End If
        
    Wend
    get_textbox_number_of_lines = -1
End Function


Function get_textbox_number_of_lines(heights() As Double, shape As Variant) ' ternary finder
    Dim line_spacing As Integer
    Dim font_size As Integer
    
    target_height = shape.Height
    line_spacing = get_i_of_line_spacing(shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin)
    font_size = get_i_of_font(shape.TextFrame.TextRange.font.Size)
    
    If font_size < 1 Then ' lower bound on font
        shape.TextFrame.TextRange.font.Size = get_font_size_i(1)
        font_size = 1
    ElseIf font_size > get_number_of_fonts() Then ' upper bound on font
        shape.TextFrame.TextRange.font.Size = get_font_size_i(get_number_of_fonts())
        font_size = get_number_of_fonts()
    End If
    
    last_index = -1
    actual_length = get_utf_string_length(shape.TextFrame.TextRange.text)
    For i = actual_length To 1 Step -1
       If Asc(Mid(shape.TextFrame.TextRange, i, 1)) > 33 And last_index = -1 Then ' finding the last character index
           last_index = i + 1
       End If
    Next i
                
    shape.TextFrame.TextRange.Characters(actual_length, 1).text = shape.TextFrame.TextRange.Characters(actual_length, 1).text & Chr(11)
    shape.TextFrame.TextRange.Characters(last_index, 1).font.Size = 1
                     
    ternary_search_left = 1
    ternary_search_right = get_maximum_number_of_lines()
    While ternary_search_left < ternary_search_right
        third = Int((ternary_search_right - ternary_search_left) / 3)
        leftThird = ternary_search_left + third
        rightThird = ternary_search_right - third
        
        cost_leftThird = Abs(heights(line_spacing, leftThird, font_size) - target_height)
        cost_rightThird = Abs(heights(line_spacing, rightThird, font_size) - target_height)
        
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
    get_textbox_number_of_lines = ternary_search_left
End Function

