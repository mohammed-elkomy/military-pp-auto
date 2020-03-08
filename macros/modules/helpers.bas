Attribute VB_Name = "helpers"
Function get_ungroupedshapes(shaperange As Variant) As Variant
    Set linear_shapes = get_linear_shapes(shaperange)
    
    For Each oshp In linear_shapes
        temp_height = oshp.Height
        oshp.Cut
        shaperange.Paste
        shaperange.Item(shaperange.Count).Height = temp_height
    Next oshp
    
    Set get_ungroupedshapes = shaperange
End Function


Function shaperange_correction()
    Set shapes_names = get_selected_names()
    fix_office_bug ActiveWindow.View.slide
    ActiveWindow.View.slide.shapes.Range(shapes_names.toArray()).Select
End Function


Function get_sub_ungroupedshapes(shaperange As Variant, oshp As Variant) As Variant
    Set linear_shapes = CreateObject("System.Collections.ArrayList")
    
    add_linear_shape oshp, linear_shapes
    
    Set shapes_names = CreateObject("System.Collections.ArrayList")
    For Each oshp In linear_shapes
        shapes_names.Add oshp.Name
        oshp.Cut
        shaperange.Paste
    Next oshp
    
    Set get_sub_ungroupedshapes = shaperange.Range(shapes_names.toArray())
End Function


Function is_arabic(chars As TextRange)
    is_arabic = (chars.LanguageID = msoLanguageIDArabicYemen Or _
    chars.LanguageID = msoLanguageIDArabicUAE Or _
    chars.LanguageID = msoLanguageIDArabicTunisia Or _
    chars.LanguageID = msoLanguageIDArabicSyria Or _
    chars.LanguageID = msoLanguageIDArabicQatar Or _
    chars.LanguageID = msoLanguageIDArabicOman Or _
    chars.LanguageID = msoLanguageIDArabicMorocco Or _
    chars.LanguageID = msoLanguageIDArabicLibya Or _
    chars.LanguageID = msoLanguageIDArabicLebanon Or _
    chars.LanguageID = msoLanguageIDArabicKuwait Or _
    chars.LanguageID = msoLanguageIDArabicJordan Or _
    chars.LanguageID = msoLanguageIDArabicIraq Or _
    chars.LanguageID = msoLanguageIDArabicJordan Or _
    chars.LanguageID = msoLanguageIDArabicEgypt Or _
    chars.LanguageID = msoLanguageIDArabicBahrain Or _
    chars.LanguageID = msoLanguageIDArabicAlgeria Or _
    chars.LanguageID = msoLanguageIDArabic)
End Function



Function get_shape_color_type(shape)

    Set Index = get_colors_inverted_index()
    If shape.Fill.GradientVariant <> 0 Then
        get_shape_color_type = Index(Str(shape.Fill.ForeColor.rgb) & "g")
    Else
        get_shape_color_type = Index(Str(shape.Fill.ForeColor.rgb))
    End If
End Function

'Function is_title_text_box(oshp As Variant)
'    is_title_text_box = oshp.Fill.GradientVariant <> 0 Or oshp.Fill.ForeColor.rgb = 255
'End Function

Function nearest_width(shape As Variant)
    w1 = 0.94 * shape.Parent.CustomLayout.Width
    w2 = (7 / 8) * shape.Parent.CustomLayout.Width
    w3 = (6 / 8) * shape.Parent.CustomLayout.Width
    w4 = (5 / 8) * shape.Parent.CustomLayout.Width
    w5 = (4 / 8) * shape.Parent.CustomLayout.Width
    w_mid_5_1 = 0.47 * shape.Parent.CustomLayout.Width
    w6 = (3 / 8) * shape.Parent.CustomLayout.Width
    w7 = (2 / 8) * shape.Parent.CustomLayout.Width
    w8 = (1 / 8) * shape.Parent.CustomLayout.Width
    
    c1 = (shape.Width - w1) * (shape.Width - w1)
    c2 = (shape.Width - w2) * (shape.Width - w2)
    c3 = (shape.Width - w3) * (shape.Width - w3)
    c4 = (shape.Width - w4) * (shape.Width - w4)
    c5 = (shape.Width - w5) * (shape.Width - w5)
    c_mid_5_1 = (shape.Width - w_mid_5_1) * (shape.Width - w_mid_5_1)
    c6 = (shape.Width - w6) * (shape.Width - w6)
    c7 = (shape.Width - w7) * (shape.Width - w7)
    c8 = (shape.Width - w8) * (shape.Width - w8)
   
    
    If c1 <= c2 And c1 <= c3 And c1 <= c4 And c1 <= c5 And c1 <= c_mid_5_1 And c1 <= c6 And c1 <= c7 And c1 <= c8 Then
        nearest_width = w1
    ElseIf c2 <= c1 And c2 <= c3 And c2 <= c4 And c2 <= c5 And c2 <= c_mid_5_1 And c2 <= c6 And c2 <= c7 And c2 <= c8 Then
        nearest_width = w2
    ElseIf c3 <= c1 And c3 <= c2 And c3 <= c4 And c3 <= c5 And c3 <= c_mid_5_1 And c3 <= c6 And c3 <= c7 And c3 <= c8 Then
        nearest_width = w3
    ElseIf c4 <= c1 And c4 <= c2 And c4 <= c3 And c4 <= c5 And c4 <= c_mid_5_1 And c4 <= c6 And c4 <= c7 And c4 <= c8 Then
        nearest_width = w4
    ElseIf c5 <= c1 And c5 <= c2 And c5 <= c3 And c5 <= c4 And c5 <= c_mid_5_1 And c5 <= c6 And c5 <= c7 And c5 <= c8 Then
        nearest_width = w5
     ElseIf c_mid_5_1 <= c1 And c_mid_5_1 <= c2 And c_mid_5_1 <= c3 And c_mid_5_1 <= c4 And c_mid_5_1 <= c5 And c_mid_5_1 <= c6 And c_mid_5_1 <= c7 And c_mid_5_1 <= c8 Then
        nearest_width = w_mid_5_1
    ElseIf c6 <= c1 And c6 <= c2 And c6 <= c3 And c6 <= c4 And c6 <= c5 And c6 <= c_mid_5_1 And c6 <= c7 And c6 <= c8 Then
        nearest_width = w6
    ElseIf c7 <= c1 And c7 <= c2 And c7 <= c3 And c7 <= c4 And c7 <= c5 And c7 <= c_mid_5_1 And c7 <= c6 And c7 <= c8 Then
        nearest_width = w7
    Else
        nearest_width = w8
    End If
    
End Function

Function get_next_width(shape As Variant)
   w1 = 0.94 * shape.Parent.CustomLayout.Width
    w2 = (7 / 8) * shape.Parent.CustomLayout.Width
    w3 = (6 / 8) * shape.Parent.CustomLayout.Width
    w4 = (5 / 8) * shape.Parent.CustomLayout.Width
    w5 = (4 / 8) * shape.Parent.CustomLayout.Width
    w_mid_5_1 = 0.47 * shape.Parent.CustomLayout.Width
    w6 = (3 / 8) * shape.Parent.CustomLayout.Width
    w7 = (2 / 8) * shape.Parent.CustomLayout.Width
    w8 = (1 / 8) * shape.Parent.CustomLayout.Width
    
    c1 = (shape.Width - w1) * (shape.Width - w1)
    c2 = (shape.Width - w2) * (shape.Width - w2)
    c3 = (shape.Width - w3) * (shape.Width - w3)
    c4 = (shape.Width - w4) * (shape.Width - w4)
    c5 = (shape.Width - w5) * (shape.Width - w5)
    c_mid_5_1 = (shape.Width - w_mid_5_1) * (shape.Width - w_mid_5_1)
    c6 = (shape.Width - w6) * (shape.Width - w6)
    c7 = (shape.Width - w7) * (shape.Width - w7)
    c8 = (shape.Width - w8) * (shape.Width - w8)
   

    If c3 <= c1 And c3 <= c2 And c3 <= c4 And c3 <= c5 And c3 <= w_mid_5_1 And c3 <= c6 And c3 <= c7 And c3 <= c8 Then
        get_next_width = w2
    ElseIf c4 <= c1 And c4 <= c2 And c4 <= c3 And c4 <= c5 And c4 <= w_mid_5_1 And c4 <= c6 And c4 <= c7 And c4 <= c8 Then
        get_next_width = w3
    ElseIf c5 <= c1 And c5 <= c2 And c5 <= c3 And c5 <= c4 And c5 <= w_mid_5_1 And c5 <= c6 And c5 <= c7 And c5 <= c8 Then
        get_next_width = w4
     ElseIf c_mid_5_1 <= c1 And c_mid_5_1 <= c2 And c_mid_5_1 <= c3 And c_mid_5_1 <= c4 And c_mid_5_1 <= c5 And c_mid_5_1 <= c6 And c_mid_5_1 <= c7 And c_mid_5_1 <= c8 Then
        get_next_width = w5
    ElseIf c6 <= c1 And c6 <= c2 And c6 <= c3 And c6 <= c4 And c6 <= c5 And c6 <= w_mid_5_1 And c6 <= c7 And c6 <= c8 Then
        get_next_width = w_mid_5_1
    ElseIf c7 <= c1 And c7 <= c2 And c7 <= c3 And c7 <= c4 And c7 <= c5 And c7 <= w_mid_5_1 And c7 <= c6 And c7 <= c8 Then
        get_next_width = w6
    ElseIf c8 <= c1 And c8 <= c2 And c8 <= c3 And c8 <= c4 And c8 <= c5 And c8 <= w_mid_5_1 And c8 <= c6 And c8 <= c7 Then
        get_next_width = w8
    Else
        get_next_width = w1
    End If
    
End Function


Function get_slide_type(slide As Variant) As String
    ' table ,simple textbox, grid of textboxes , image,title slide
    
    slide_type = "other"
    For Each oshp In slide.shapes
        shape_type = get_shape_type(oshp)
        If shape_type = "picture" Or shape_type = "table" Or shape_type = "group" Or shape_type = "other" Then
            slide_type = shape_type
            get_slide_type = slide_type
            Exit Function
        End If
    Next oshp
    
    If slide.shapes.Count() = 1 Then
        slide_type = "title slide"
    Else
        slide_type = "simple textbox"
    End If
    
    get_slide_type = slide_type
End Function



Function get_shape_type(shape As Variant) As String
    ' assuming primitive shapes
    ' table ,textbox, picture,other(contains arrows drawn and most drawn things and cliparts) , empty
    
    shape_type = "other"
    If shape.Type = msoAutoShape And shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            shape_type = "textbox"
        ElseIf shape.AutoShapeType = msoShapeRoundedRectangle Then
            shape_type = "empty"
        End If
    ElseIf shape.Type = msoTable Then
        shape_type = "table"
    ElseIf shape.Type = msoPicture Then
        shape_type = "picture"
    ElseIf shape.Type = msoGroup Then
        shape_type = "group"
    ElseIf shape.Type = msoTextBox Then
        shape_type = "textbox"
    End If
    
    get_shape_type = shape_type
End Function




Function unique(original As Variant) As Variant
    Set unqiue_alist = CreateObject("System.Collections.ArrayList")
    For Each original_element In original
        found = False
        For Each unqiue_element In unqiue_alist
            If unqiue_element = original_element Then
                found = True
            End If
        Next unqiue_element
        
        If found = False Then
            unqiue_alist.Add original_element
        End If
    Next original_element
    
    Set unique = unqiue_alist
End Function

Function deg_to_rad(deg As Double) As Double
    pi = 4 * Atn(1)
    deg_to_rad = deg / 180 * pi
End Function

Function max(first As Double, second As Double) As Double
    If first > second Then
        max = first
    Else
        max = second
    End If
End Function

Function min(first As Double, second As Double) As Double
    If first < second Then
        min = first
    Else
        min = second
    End If
End Function

Function min_int(first As Integer, second As Integer) As Integer
    If first < second Then
        min_int = first
    Else
        min_int = second
    End If
End Function


Function string_to_rgb(rgb_str As Variant)
    Dim dataArray
    dataArray = Split(rgb_str, ",", 3)
    string_to_rgb = rgb(Int(dataArray(0)), Int(dataArray(1)), Int(dataArray(2)))
End Function



Function rgb_div(rgb_tuple As Variant) As Variant
    Dim temp As Long
    temp = rgb_tuple
    Set rgb_lst = unrgb(temp)
    rgb_div = rgb(Int(rgb_lst(0) / 2.161), Int(rgb_lst(1) / 2.161), Int(rgb_lst(2) / 2.161))
End Function

Function rgb(r As Integer, g As Integer, b As Integer) As Long
    Dim base As Long
    base = 256
    rgb = base * base * b + base * g + r
End Function


Function unrgb(rgb_long As Long) As Variant
    Dim base As Long
    Dim rg As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long

    base = 256
    
    b = Int(rgb_long / (base * base))
    rg = (rgb_long - b * base * base)
    g = Int(rg / base)
    r = rg - g * base
    
    Set return_val = CreateObject("System.Collections.ArrayList")
    return_val.Add (r)
    return_val.Add (g)
    return_val.Add (b)
    
    Set unrgb = return_val
End Function

Function rgb_distance(first As Long, second As Long) As Long
    Set first_rgb = unrgb(first)
    Set second_rgb = unrgb(second)
    wieghts = Array(0.3, 0.59, 0.11)
    r_mean = (first_rgb(0) + second_rgb(0)) / 2
    r_sqrd = (first_rgb(0) - second_rgb(0)) * (first_rgb(0) - second_rgb(0))
    g_sqrd = (first_rgb(1) - second_rgb(1)) * (first_rgb(1) - second_rgb(1))
    b_sqrd = (first_rgb(2) - second_rgb(2)) * (first_rgb(2) - second_rgb(2))
    Dim error As Long
    error = (2 + r_mean / 256) * r_sqrd
    error = error + 4 * g_sqrd
    error = error + (2 + (255 - r_mean) / 256) * b_sqrd
    
    rgb_distance = error
End Function

Function get_utf_string_length(query_string As String)
    length = 1
    On Error GoTo Finish:

    While Len(Mid(query_string, length, 1)) <> 0
        length = length + 1
    Wend

Finish:
    get_utf_string_length = length - 1
End Function


Function group_sorted_list_archived(sorted_array As Variant)
    Dim ret_list As Object
    Set ret_list = CreateObject("System.Collections.ArrayList")
    
    Dim sublist As Object
    Set sublist = CreateObject("System.Collections.ArrayList")
    
    Dim group_bbox As vertical_projection_bbox
    Set group_bbox = New vertical_projection_bbox

    
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
    ret_list.Add sublist
    Set group_sorted_list = ret_list
End Function


Function group_sorted_list(sorted_array As Variant)
    Dim ret_list As Object
    Set ret_list = CreateObject("System.Collections.ArrayList")
    
    Dim sublist As Object
    Set sublist = CreateObject("System.Collections.ArrayList")
    
    Dim group_bbox As vertical_projection_bbox
    Set group_bbox = New vertical_projection_bbox

    
    For i = 0 To sorted_array.Count() - 1:
        Set sorted_array.Item(i).shape_bbox = get_shape_vertical_projection(sorted_array.Item(i).shape)
        If sublist.Count() > 0 Then
            If group_bbox.vertical_jaccard(sorted_array.Item(i).shape_bbox) < 0.6 And sorted_array.Item(i).shape_bbox.vertical_jaccard(group_bbox) < 0.6 Then
                ret_list.Add sublist
                Set sublist = CreateObject("System.Collections.ArrayList")
                group_bbox.min_y = sorted_array.Item(i).shape_bbox.min_y
                group_bbox.max_y = sorted_array.Item(i).shape_bbox.max_y
            End If
        Else ' first item
            group_bbox.min_y = sorted_array.Item(i).shape_bbox.min_y
            group_bbox.max_y = sorted_array.Item(i).shape_bbox.max_y
        End If
        sublist.Add sorted_array.Item(i)
        
        group_bbox.min_y = min(sorted_array.Item(i).shape_bbox.min_y, group_bbox.min_y)
        group_bbox.max_y = max(sorted_array.Item(i).shape_bbox.max_y, group_bbox.max_y)
    Next i
    ret_list.Add sublist
    Set group_sorted_list = ret_list
End Function

Function get_selected_names()
    Set shapes_names = CreateObject("System.Collections.ArrayList")
    For Each shape In ActiveWindow.Selection.shaperange
            shapes_names.Add shape.Name
    Next shape
    Set get_selected_names = shapes_names
End Function

Function get_sorted_shapes_groups(array_to_sort As Variant, shapes As Variant) ' array  list expected
    sort_shapes array_to_sort
    'array_to_sort now sorted
    
    On Error GoTo nogroup:
    
    ' collections are groups of objects based on vertical intersection
    Set collections = group_sorted_list(array_to_sort)
    
    ' grouping collections using powerpoint API
    For Each Collection In collections
        Dim id_list As Object
        Set shapes_names = CreateObject("System.Collections.ArrayList")
        For Each ex_shape In Collection
            shapes_names.Add ex_shape.shape.Name
        Next ex_shape
        
        If shapes_names.Count > 1 Then
            shapes.Range(shapes_names.toArray()).Group
          
        End If
    Next Collection
    
    
    Set shapes_arraylist = CreateObject("System.Collections.ArrayList")
     'now add all of those groups
     For Each oshp In shapes
        Set extended = New extended_shape
        Set extended.shape = oshp
        Set extended.shape_bbox = get_shape_bbox(oshp)
        extended.shape_name = oshp.Name
        shapes_arraylist.Add extended
     Next oshp
     
     'sort the groups because they are given out of order from powerpoint
     sort_shapes shapes_arraylist
     Set get_sorted_shapes_groups = shapes_arraylist
     Exit Function
nogroup:
     Set get_sorted_shapes_groups = array_to_sort
End Function



Function sort_shapes(array_to_sort As Variant) ' array  list expected
    Dim X As Long, y As Long
    
    For X = 0 To array_to_sort.Count - 2
        ' find minimum in subarray
        minimum_index = X
        For y = X + 1 To array_to_sort.Count - 1
            If array_to_sort.Item(minimum_index).shape_bbox.min_y > array_to_sort.Item(y).shape_bbox.min_y Then
               minimum_index = y
            End If
        Next y
        
        ' swap x with the minimum
        Set temp = array_to_sort.Item(X)
        Set array_to_sort.Item(X) = array_to_sort.Item(minimum_index)
        Set array_to_sort.Item(minimum_index) = temp
    Next X
End Function


Function get_char_unicode(char As String) As String
    code = Hex(AscW(char))
    While Len(code) < 4
        code = "0" & code
    Wend
    
    code = "\u" & code
    get_char_unicode = code
End Function

Function get_string_unicode(sequence As String) As String
    Dim ret_str As String
    ret_str = ""
    For Index = 1 To Len(sequence)
        ret_str = ret_str & get_char_unicode(Mid(sequence, Index, 1))
    Next Index
    
    get_string_unicode = ret_str
End Function




Function hasGroupItems(shape As Variant)
    On Error GoTo cant:
        Set r = shape.GroupItems
        hasGroupItems = True
        Exit Function
cant:
    hasGroupItems = False
End Function

Function add_linear_shape(shape As Variant, linear_shapes As Variant)
    If hasGroupItems(shape) Then
        For Each groupshape In shape.GroupItems
                add_linear_shape groupshape, linear_shapes
        Next groupshape
    Else
        shape.Name = "" & shape.Id
        linear_shapes.Add (shape)
    End If
End Function

Function get_linear_shapes(shapes As Variant) As Variant
    Set linear_shapes = CreateObject("System.Collections.ArrayList")
    
    For Each shape In shapes
        add_linear_shape shape, linear_shapes
    Next shape
    
    Set get_linear_shapes = linear_shapes
End Function


Function make_regex(pattern As String)
    Set reg = CreateObject("vbscript.regexp")
    With reg
    .Global = True
    .pattern = pattern
    End With
    Set make_regex = reg
End Function

Function fix_office_bug(slide As Variant)
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
    slide.shapes.SelectAll
    ActiveWindow.Selection.Cut
    slide.shapes.Paste
End Function


Function is_correct_shape(shape As Variant)
    On Error GoTo cant:
        get_shape_type shape
        is_correct_shape = True
        Exit Function
cant:
    is_correct_shape = False
End Function




