Attribute VB_Name = "global_variables"
' global variables
Public heights() As Double

Dim light_color_AL As Variant
Dim dark_color_AL As Variant
Dim dark_grad_color_AL As Variant
Dim light_grad_color_AL As Variant

Dim colors_inverted_index As Variant

Dim table_border_rgb As Long
Dim light_border_rgb As Long
Dim dark_border_rgb As Long
Dim light_gardient_border_rgb As Long
Dim dark_gardient_border_rgb As Long
Dim image_border_rgb As Long

Dim header_font_size As Long
Dim inner_font_size As Long
Dim main_title_font_size As Long

Dim title_space_multiplier As Double
' end: global variables

' global variables getters
Function get_heights()
    make_initalized
    get_heights = heights
End Function

Function get_light_color_AL() As Variant
    make_initalized
    Set get_light_color_AL = light_color_AL
End Function

Function get_dark_color_AL() As Variant
    make_initalized
    Set get_dark_color_AL = dark_color_AL
End Function

Function get_dark_grad_color_AL() As Variant
    make_initalized
    Set get_dark_grad_color_AL = dark_grad_color_AL
End Function

Function get_light_grad_color_AL() As Variant
    make_initalized
    Set get_light_grad_color_AL = light_grad_color_AL
End Function

Function get_table_border_rgb() As Long
    make_initalized
    get_table_border_rgb = table_border_rgb
End Function

Function get_light_border_rgb() As Long
   make_initalized
   get_light_border_rgb = light_border_rgb
End Function

Function get_dark_border_rgb() As Long
    make_initalized
    get_dark_border_rgb = dark_border_rgb
End Function

Function get_dark_gardient_border_rgb() As Long
    make_initalized
    get_dark_gardient_border_rgb = dark_gardient_border_rgb
End Function

Function get_light_gardient_border_rgb() As Long
    make_initalized
    get_light_gardient_border_rgb = light_gardient_border_rgb
End Function

Function get_image_border_rgb() As Long
    make_initalized
    get_image_border_rgb = image_border_rgb
End Function


Function get_inner_font_size() As Long
    make_initalized
    get_inner_font_size = inner_font_size
End Function

Function get_header_font_size() As Long
    make_initalized
    get_header_font_size = header_font_size
End Function

Function get_main_title_font_size() As Long
    make_initalized
    get_main_title_font_size = main_title_font_size
End Function

Function get_colors_inverted_index()
    make_initalized
    Set get_colors_inverted_index = colors_inverted_index
End Function

Function get_title_space_multiplier() As Double
    make_initalized
    get_title_space_multiplier = title_space_multiplier
End Function

'end: global variables getters

' constants getters
Function get_line_spacing_precision() As Integer
    get_line_spacing_precision = 300 'line spacing precision
End Function


Function get_root_dir() As String
    If Application.AddIns.Count > 0 Then
        get_root_dir = Application.AddIns.Item(1).Path & "\"
    Else
        get_root_dir = ActivePresentation.Path & "\"
    End If
End Function

Function get_maximum_number_of_lines() As Integer
    get_maximum_number_of_lines = 25 'max number of lines
End Function

Function get_number_of_fonts() As Integer
    get_number_of_fonts = 32 - 10 + 1 'number of fonts 10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32
End Function

Function get_lines_ub() As Double
    ' maximum line spacing to be used
    get_lines_ub = 3.2
End Function

Function get_lines_lb() As Double
    ' minimum line spacing to be used
    get_lines_lb = 0.6
End Function

Function get_inner_light_colors()
    get_inner_light_colors = get_light_color_AL().toArray()
End Function


Function get_inner_dark_colors()
    get_inner_dark_colors = get_dark_color_AL().toArray()
End Function

Function get_dark_gradient_colors()
     get_dark_gradient_colors = get_dark_grad_color_AL().toArray()
End Function

Function get_light_gradient_colors()
     get_light_gradient_colors = get_light_grad_color_AL().toArray()
End Function

Function get_text_colors()
    Dim colours(1 To 9) As Long
    ' colors list for text
    colours(1) = rgb(255, 255, 255)
    colours(2) = rgb(0, 0, 0)
    colours(3) = rgb(0, 0, 255)
    colours(4) = rgb(0, 102, 0)
    colours(5) = rgb(255, 0, 0)
    colours(6) = rgb(255, 255, 0)
    colours(7) = rgb(255, 204, 255)
    colours(8) = rgb(0, 255, 0)
    colours(9) = rgb(0, 255, 255)
    
    get_text_colors = colours
End Function

Function get_header_text_colors()
    Dim colours(1 To 6) As Long
    ' colors list for text
    colours(1) = rgb(255, 255, 255)
    colours(2) = rgb(255, 255, 0)
    colours(3) = rgb(255, 255, 0)
    colours(4) = rgb(0, 255, 0)
    colours(5) = rgb(0, 0, 0)
    colours(6) = rgb(0, 0, 0)
    
    get_header_text_colors = colours
End Function


Function get_numbers_and_dates_regex()
    Set regX = CreateObject("vbscript.regexp")
    
    With regX
    .Global = True
    .pattern = "\d((\d|\\|\u002d|\u002f|/|\s|\u2013)*)(\d|\u0645(\s|$)|%)|\d" ' regex for dates and numbers
    End With
    
    Set get_numbers_and_dates_regex = regX
End Function




Sub make_initalized()
    On Error GoTo initailize
        temp = UBound(heights)
        GoTo skip
initailize:
        heights = read_heights_file()
        parse_configs
        build_colors_inverted_index
skip:

End Sub

Sub force_initalize()
    heights = read_heights_file()
    parse_configs
    build_colors_inverted_index

End Sub

Sub build_colors_inverted_index()
    local_light_colors = get_inner_light_colors()
    local_dark_colors = get_inner_dark_colors()
    local_dark_gradient_colors = get_dark_gradient_colors()
    local_light_gradient_colors = get_light_gradient_colors()
    Set colors_inverted_index = CreateObject("System.Collections.Hashtable")
    
    For i = LBound(local_light_colors) To UBound(local_light_colors)
        If Not colors_inverted_index.containskey(Str(local_light_colors(i))) Then
             colors_inverted_index.Add Str(local_light_colors(i)), "light"
        End If
    Next i
    
    For i = LBound(local_dark_colors) To UBound(local_dark_colors)
        If Not colors_inverted_index.containskey(Str(local_dark_colors(i))) Then
            colors_inverted_index.Add Str(local_dark_colors(i)), "dark"
        End If
        
        
    Next i
    
    For i = LBound(local_light_gradient_colors) To UBound(local_light_gradient_colors)
        If Not colors_inverted_index.containskey(Str(rgb_div(local_light_gradient_colors(i))) & "g") Then
            colors_inverted_index.Add Str(rgb_div(local_light_gradient_colors(i))) & "g", "light"
        End If
        
    Next i
    
    For i = LBound(local_dark_gradient_colors) To UBound(local_dark_gradient_colors)
        If Not colors_inverted_index.containskey(Str(rgb_div(local_dark_gradient_colors(i))) & "g") Then
            colors_inverted_index.Add Str(rgb_div(local_dark_gradient_colors(i))) & "g", "dark"
        End If
    Next i
    
End Sub

Sub parse_configs()
    Set configs = ParseYAML()
    Set light_color_AL = CreateObject("System.Collections.ArrayList")
    Set dark_color_AL = CreateObject("System.Collections.ArrayList")
    Set dark_grad_color_AL = CreateObject("System.Collections.ArrayList")
    Set light_grad_color_AL = CreateObject("System.Collections.ArrayList")
    
    table_border_rgb = string_to_rgb(configs("#borders")("table"))
    light_border_rgb = string_to_rgb(configs("#borders")("light"))
    dark_border_rgb = string_to_rgb(configs("#borders")("dark"))
    light_gardient_border_rgb = string_to_rgb(configs("#borders")("light-gradient"))
    dark_gardient_border_rgb = string_to_rgb(configs("#borders")("dark-gradient"))
    
    image_border_rgb = string_to_rgb(configs("#borders")("image"))
    
    
    header_font_size = Int(configs("#font_size")("header"))
    inner_font_size = Int(configs("#font_size")("inner"))
    main_title_font_size = Int(configs("#font_size")("main_title"))
    title_space_multiplier = CDbl(configs("#format")("title_space_multiplier"))
    
    
    For Each light_color In configs("#light")
        light_color_AL.Add string_to_rgb(light_color)
    Next light_color

    
    For Each dark_color In configs("#dark")
        dark_color_AL.Add string_to_rgb(dark_color)
    Next dark_color
    
    For Each dark_grad_color In configs("#dark_grad")
        dark_grad_color_AL.Add string_to_rgb(dark_grad_color)
    Next dark_grad_color
    
    For Each light_grad_color In configs("#light_grad")
        light_grad_color_AL.Add string_to_rgb(light_grad_color)
    Next light_grad_color
    
End Sub






