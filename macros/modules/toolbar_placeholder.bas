Attribute VB_Name = "toolbar_placeholder"
Sub placeholder_toolbar(toolbarname As String)
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
  
    'on error go to
    On Error GoTo escape
        CommandBars(toolbarname).Delete
escape:

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=toolbarname, _
        Position:=msoBarRight, Temporary:=False)
        
    Set optimal_finder_button = oToolbar.Controls.Add(Type:=msoControlButton)
    optimal_finder_button.Enabled = False
  
    
    oToolbar.Visible = True
End Sub















