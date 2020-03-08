VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vertical_projection_bbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' one dimensional intersection helper

Public min_y As Double
Public max_y As Double

Function length() As Double
    length = (max_y - min_y)
End Function

Function intersect(that As vertical_projection_bbox) As Double
    If is_disjoint(that) Then
        intersect = -1
    Else
    
        Set intersection = New vertical_projection_bbox
        intersection.min_y = max(that.min_y, min_y)
        intersection.max_y = min(that.max_y, max_y)
        intersect = intersection.length()
    End If
End Function

Function union(that As vertical_projection_bbox) As Double
    If is_disjoint(that) Then
        union = -1
    Else
    
        Set union_ = New vertical_projection_bbox
        union_.min_y = min(that.min_y, min_y)
        union_.max_y = max(that.max_y, max_y)
        union = union_.length()
    End If
End Function


Function vertical_jaccard(that As vertical_projection_bbox) As Double
    intersection = intersect(that)
    'union_ = union(that)
     'If union_ > 0 Then
    '    jaccard = intersection / union_
    
    this_length = length()
    that_length = that.length()
    If this_length > 0 And intersection > 0 Then
        vertical_jaccard = intersection / that_length
        'vertical_jaccard = (intersection / this_length + intersection / that_length) / 2
    Else
        vertical_jaccard = -1
    End If
End Function


Function is_disjoint(that As vertical_projection_bbox) As Boolean
    is_disjoint = that.min_y > max_y Or _
                  min_y > that.max_y
End Function





