Attribute VB_Name = "Helpers_gui"
Function make_progressor()
    progressor.Show vbModeless
    progressor.status.Width = 1
End Function

Function done_progressor()
    progressor.hide
End Function

Function update_progressor(val)
    Dim temp As Double
    temp = val
    progressor.status.Width = max(0, min(temp, 100))
End Function


Function make_animation_range()
    animation.Show vbModeless
End Function

Function make_un_animation_range()
    un_animation.Show vbModeless
End Function

Function make_grid_size()
    grid_creator.Show vbModeless
End Function


Sub TestTheForm()
    make_progressor
    
    For i = -50 To 120
        For j = 1 To 1000000
        Next j
        update_progressor i
        DoEvents
    Next i
    
End Sub

Sub main()
    
    

    
    
 For Each cbar In CommandBars
    Debug.Print cbar.Name
    For Each Control In cbar.Controls
        Debug.Print Control.Caption
    Next Control
    Debug.Print "--"
 Next cbar
    

End Sub


Sub disble_bar(toolbarname)
    For Each Control In CommandBars(toolbarname).Controls
        Control.Enabled = False
    Next Control
End Sub

Sub enable_bar(toolbarname)
    For Each Control In CommandBars(toolbarname).Controls
        Control.Enabled = True
    Next Control
End Sub


Function show_error(message As String)
    MsgBox message, vbCritical
End Function


