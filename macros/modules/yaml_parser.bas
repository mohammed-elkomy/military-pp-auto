Attribute VB_Name = "yaml_parser"

Function ParseYAML()
    Dim myFile As String, text As String, textline As String
    ' open YAML file
   
    Open get_root_dir & "config.komy.txt" For Input As #1
    Dim dataArray
    Dim last_key As String
    Dim configs As Object
    Set configs = CreateObject("System.Collections.Hashtable")

    Dim Group As Collection
    Set Group = New Collection
    
    
    Line = 0
    Do Until EOF(1)
        Line Input #1, textline
        oneline = Replace(textline, " ", "")
        dataArray = Split(oneline, ":", 2)
        sizeArray = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
        ' Verification Empty Lines and Split don't occur
        If Not textline = "" And Not sizeArray = 0 Then
            Data = dataArray(1)
            key = dataArray(0)
            ' test if line don't start with -
            If InStr(1, key, "#") = 0 Then
                Group.Add Data, key
            Else:
                If last_key <> "" Then
                        configs.Add last_key, Group
                        Set Group = New Collection
                End If
                last_key = key
            End If
            ' just for debug
            Line = Line + 1
            'text = text & textline
        End If
    Loop
    Close #1

    If last_key <> "" Then
        configs.Add last_key, Group
        Set Group = New Collection
    End If
    
    Set ParseYAML = configs

End Function
 
 











