Attribute VB_Name = "creator_op"
Sub create_grid_hidden(rows, columns)
    Set placeholder = ActiveWindow.Selection.shaperange.Item(1)
    
    active_width = placeholder.Width
    col_padding = active_width * 0.1 / (columns - 1)
    row_padding = placeholder.Height * 0.1 / (rows - 1)
    cell_width = active_width * 0.9 / columns
   
    Dim cell_height As Double
    cell_height = 39
    do_padding = 0
    
    For r = 0 To rows - 1
        For c = 0 To columns - 1
            If c = 0 Then
                do_padding = 0
            Else
                do_padding = 1
            End If
            
            ActiveWindow.Selection.SlideRange.shapes.AddShape(msoShapeRoundedRectangle, placeholder.Left + c * cell_width + col_padding * c, placeholder.Top + r * cell_height, cell_width, cell_height).Select  'x(720),y(540),w,h
            ActiveWindow.Selection.shaperange.Shadow.Visible = False
            
            With ActiveWindow.Selection.shaperange
                .Fill.ForeColor.rgb = rgb(0, 255, 0)
                .Fill.Visible = msoTrue
                .Fill.Solid
                
                .Line.ForeColor.rgb = rgb(255, 0, 0)
                .Line.Visible = msoTrue
                .Fill.Transparency = 0#
                .Line.Weight = 3#
            End With
            
            ActiveWindow.Selection.shaperange.TextFrame.TextRange.Characters(start:=1, length:=0).Select
            With ActiveWindow.Selection.TextRange
                .text = "»‰œ" + Chr$(CharCode:=11)
                With .font
                    .Name = "Arial"
                    .NameComplexScript = "Arial"
                    .NameOther = "Arial"
                    .Size = 22
                    .Bold = msoFalse
                    .Italic = msoFalse
                    .Underline = msoFalse
                    .Shadow = msoFalse
                    .Emboss = msoFalse
                    .BaselineOffset = 0
                    .AutoRotateNumbers = msoFalse
                    .color.SchemeColor = ppForeground
                    .NameAscii = "PT Bold Heading"
                    .NameComplexScript = "PT Bold Heading"
                End With
                .ParagraphFormat.Alignment = ppAlignJustifyLow
            End With
            
            ActiveWindow.Selection.shaperange.TextFrame.TextRange.Characters(start:=4, length:=1).Select
            ActiveWindow.Selection.TextRange.font.Size = 1
            ActiveWindow.Selection.shaperange.Top = placeholder.Top + r * cell_height + row_padding * r
        Next c
    Next r
End Sub










