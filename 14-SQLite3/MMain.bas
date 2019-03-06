Attribute VB_Name = "MMain"
Option Explicit

Enum ReturnCode
    ErrRT = -1
    SuccessRT = 1
End Enum

Enum Pos
    RowStart = 2
End Enum

Type FieldInfo
    FName As String
    Type As String
    pk As Long '主键，0不是 1是
End Type

Type TableInfo
    tableName As String
    Fields() As FieldInfo
End Type

Type DataStruct
    Sqlite As CSQLite3
        
End Type

Sub vba_main()
    Dim d As DataStruct
    
    Set d.Sqlite = New CSQLite3
    #If Win64 Then
        If d.Sqlite.Initialize(ThisWorkbook.Path + "\x64") = d.Sqlite.InitERR Then Exit Sub
    #Else
        If d.Sqlite.Initialize() = d.Sqlite.InitERR Then Exit Sub
    #End If
    
    Debug.Print "Version=", d.Sqlite.Version

    d.Sqlite.SetDBName = ThisWorkbook.Path & "\price.db"
    Dim ret As Long
    ret = d.Sqlite.OpenDB()
    If ret Then
        Debug.Print d.Sqlite.GetErr()
    End If
    
    
    Dim i As Long, j As Long
    Dim arr() As TableInfo
    ret = d.Sqlite.GetTableInfo(arr)
    If ret Then
        Debug.Print d.Sqlite.GetErr()
    Else
        For i = 0 To UBound(arr)
            AddSht arr(i).tableName
            Cells(1, 1).Value = "序号"
            Cells(1, 2).Value = "FieldName"
            Cells(1, 3).Value = "FieldType"
            Cells(1, 4).Value = "主键"

            For j = 0 To UBound(arr(i).Fields)
                Cells(j + 2, 1).Value = j + 1
                Cells(j + 2, 2).Value = arr(i).Fields(j).FName
                Cells(j + 2, 3).Value = arr(i).Fields(j).Type
                Cells(j + 2, 4).Value = arr(i).Fields(j).pk
            Next
        Next
    End If
    
    ret = d.Sqlite.CloseDB()
    If ret Then
        Debug.Print d.Sqlite.GetErr()
    End If
    
    On Error GoTo err_handle
    
    Set d.Sqlite = Nothing
    Exit Sub
err_handle:
    MsgBox Err.Description
End Sub

Function InsertMore(d As DataStruct) As Long
    Dim ret As Long
    
    ret = d.Sqlite.BeginTransaction("INSERT INTO testTable Values (?, ?)")
    If ret Then
        InsertMore = ret
        Exit Function
    End If
    
    Dim i As Long
    For i = 10 To 20
        ret = d.Sqlite.BindInt32(1, i)
        If ret Then
            InsertMore = ret
            Exit Function
        End If
        
        ret = d.Sqlite.BindText(2, "abc" & VBA.CStr(i))
        If ret Then
            InsertMore = ret
            Exit Function
        End If
        
        ret = d.Sqlite.Step()
        If ret <> d.Sqlite.StepDone Then
            InsertMore = ret
            Exit Function
        End If
    
        ret = d.Sqlite.Reset()
        If ret Then
            InsertMore = ret
            Exit Function
        End If
        
    Next
    
    InsertMore = d.Sqlite.CommitTransaction()
End Function

Function AddSht(sht_name As String)
    On Error Resume Next
    ActiveWorkbook.Worksheets(sht_name).Activate
    If Err.Number <> 0 Then
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = sht_name
    Else
        Cells.Delete
    End If
    On Error GoTo 0
End Function

