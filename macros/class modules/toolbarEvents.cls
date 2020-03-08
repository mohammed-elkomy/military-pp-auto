VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "toolbarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1


Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    On Error GoTo none
    
    Select Case Sel.Type
        Case ppSelectionText
            inner_colors_disable
            shaperange_disable
            text_colors_enable
            
        Case ppSelectionSlides
            inner_colors_disable
            shaperange_disable
            text_colors_disable
            If ActiveWindow.Selection.SlideRange.Count = 1 Then
                singleslide_enable
            Else
                singleslide_disable
            End If
            
            
        Case ppSelectionShapes
            inner_colors_enable
            shaperange_enable
            text_colors_enable
         
        Case ppSelectionNone
none:
            inner_colors_disable
            shaperange_disable
            text_colors_disable
    End Select
    
    
End Sub





