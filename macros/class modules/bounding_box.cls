VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bounding_box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' two dimensional intersection helper

Public min_x As Double
Public min_y As Double
Public max_x As Double
Public max_y As Double
Public center_x As Double
Public center_y As Double

Function area() As Double
    area = (max_y - min_y) * (max_x - min_x)
End Function

Function intersect(that As bounding_box) As Double
    If is_disjoint(that) Then
        intersect = -1
    Else
    
        Set intersection = New bounding_box
        
        intersection.min_x = max(that.min_x, min_x)
        intersection.min_y = max(that.min_y, min_y)
        intersection.max_x = min(that.max_x, max_x)
        intersection.max_y = min(that.max_y, max_y)
        intersect = intersection.area()
    End If
    
End Function


Function jaccard(that As bounding_box) As Double
    intersection = intersect(that)
    If that.area() > 0 Then
        jaccard = intersection / that.area()
    Else
        jaccard = -1
    End If
End Function


Function is_disjoint(that As bounding_box) As Boolean
    is_disjoint = that.min_x > max_x Or _
                  min_x > that.max_x Or _
                  that.min_y > max_y Or _
                  min_y > that.max_y
End Function



