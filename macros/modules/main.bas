Attribute VB_Name = "main"
Public X As New toolbarEvents
Sub Auto_Open()
    force_initalize
    Set X.App = Application
    On Error GoTo skip
    
    ActivePresentation.ExtraColors.Clear
    inner_colors = get_inner_light_colors()
    
    For i = UBound(inner_colors) To LBound(inner_colors) Step -1
        ActivePresentation.ExtraColors.Add inner_colors(i)
    Next
    
skip:
    
    inner_colors_toolbar
    text_toolbar
    shaperange_toolbar
    optimizer_toolbar
    grid_toolbar
    
    placeholder_toolbar ("place")
    slidebased_toolbar
    placeholder_toolbar ("place2")
    about_toolbar
    MsgBox "         . „  Õ„Ì· «·≈÷«›… »‰Ã«Õ", Title:="       ÂÌ∆…  œ—Ì» «·ﬁÊ«  «·„”·Õ…"
End Sub


Sub main()
   force_initalize
        


End Sub







