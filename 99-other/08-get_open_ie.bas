Attribute VB_Name = "模块4"
Option Explicit


'Microsoft HTML Object Library
'Microsoft Internet Controls

Sub te()
    Dim ie As Object
    Dim obj_frame As FramesCollection
    Dim hw As HTMLWindow2, sub_hw As HTMLWindow2
    Dim a
    Dim i As Long
    
    Set ie = get_open_ie("url")
    
    Set obj_frame = ie.Document.frames
    
    Set hw = obj_frame(1)
    
    Set obj_frame = hw.Document.frames
    Set sub_hw = obj_frame(0)

    With sub_hw.Document
'        标题
        .getElementById("id").Focus
        .getElementById("id").Value = "xx"
      
        '提交
        .getElementById("buttonSave").Click
    End With
    
    
    Set ie = Nothing
    Set obj_frame = Nothing
    Set hw = Nothing
End Sub

Function get_open_ie(str_url As String) As Object
    Dim obj_shell As Object
    Dim obj_ie As Object
    Dim i As Long
    
    Set obj_shell = VBA.CreateObject("Shell.Application")
    
    For i = obj_shell.Windows.Count To 1 Step -1
        Set obj_ie = obj_shell.Windows(i - 1)
        
        If obj_ie Is Nothing Then
            Exit For
        End If
        
'        Debug.Print obj_ie.FullName
        
        If VBA.Right(VBA.UCase(obj_ie.FullName), 12) = "IEXPLORE.EXE" Then
'            Debug.Print obj_ie.Document.URL
            If obj_ie.Document.URL = str_url Then
                Set get_open_ie = obj_ie
                Exit Function
            End If
        End If
    Next i
End Function

