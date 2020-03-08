Attribute VB_Name = "constructors"
Function get_shape_bbox(shape As Variant) As Variant
    ' handeling rotation equations
    ' for a shape with width:w ,height :h, left:l, top:t,rotation:r_deg
    ' find r_rad from r_deg
    ' deltaY = h/2 cos (r_rad) + w /2 sin (r_rad)
    ' deltaX = h/2 sin (r_rad) + w /2 cos (r_rad)
    ' new shape top is = center_y + abs(deltaY)
    
    Dim box As bounding_box
    Set box = New bounding_box
    On Error GoTo norotation
        rad = deg_to_rad(shape.Rotation)
        GoTo Rotation
        
norotation:
    rad = 0
Rotation:
    deltaY = Abs(shape.Height / 2 * Cos(rad) + shape.Width / 2 * Sin(rad))
    deltaX = Abs(shape.Height / 2 * Sin(rad) + shape.Width / 2 * Cos(rad))
    center_y = shape.Top + shape.Height / 2
    center_x = shape.Left + shape.Width / 2
    
    box.min_x = center_x - deltaX ' left
    box.min_y = center_y - deltaY ' top
    box.max_x = center_x + deltaX ' right
    box.max_y = center_y + deltaY ' bottom
    
    box.center_x = center_x
    box.center_y = center_y
    
    Set get_shape_bbox = box
End Function

Function get_shape_vertical_projection(shape As Variant) As Variant
    ' handeling rotation equations
    ' for a shape with height :h, top:t,rotation:r_deg
    ' find r_rad from r_deg
    ' deltaY = h/2 cos (r_rad) + w /2 sin (r_rad)
    ' new shape top is = center_y + abs(deltaY)
    
    Dim vbox As vertical_projection_bbox
    Set vbox = New vertical_projection_bbox
    rad = deg_to_rad(shape.Rotation)
    deltaY = Abs(shape.Height / 2 * Cos(rad) + shape.Width / 2 * Sin(rad))
    center_y = shape.Top + shape.Height / 2

    vbox.min_y = center_y - deltaY ' top
    vbox.max_y = center_y + deltaY ' bottom
    
    Set get_shape_vertical_projection = vbox
End Function

