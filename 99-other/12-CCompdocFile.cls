VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCompdocFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
   
Const CFHEADER_SIZE As Long = 2 ^ 9
Const DIR_SIZE As Long = 128

Private Type CFHeader
    id(7) As Byte                   '文档标识id
    file_id(15) As Byte             '文件唯一标识
    file_format_revision As Integer '文件格式修订号
    file_format_version As Integer  '文件格式版本号
    memory_endian As Integer        'FFFE表示 Little-Endian
    sector_size As Integer          '扇区的大小 2的幂 通常为2^9=512
    short_sector_size As Integer    '短扇区大小，2的幂,通常为2^6
    not_used_1(9) As Byte
    SAT_count As Long               '分区表扇区的总数
    dir_first_SID As Long           '目录流第一个扇区的ID
    not_used_2(3) As Byte
    min_stream_size As Long         '最小标准流
    SSAT_first_SID As Long          '短分区表的第一个扇区ID
    SSAT_count As Long              '短分区表扇区总数
    MSAT_first_SID As Long          '主分区表的第一个扇区ID
    MSAT_count As Long              '分区表的扇区总数
    arr_SID(108) As Long            '主分区表前109个记录
End Type

Private Type CFDir
    dir_name(63) As Byte
    len_name As Integer
    type As Byte                    '1仓storage 2流 5根
    color As Byte                    '0红色 1黑色
    left_child As Long               '-1表示叶子
    right_child As Long
    sub_dir As Long
    arr_keep(19) As Byte
    time_create As Date
    time_modify As Date
    first_SID As Long               '目录入口所表示的第1个扇区编码
    stream_size As Long             '目录入口流尺寸，可判断是否是短扇区
    not_used As Long
End Type

Private my_path As String
Private b_ready As Boolean
Dim file_byte() As Byte
Dim cf_header As CFHeader           '文件头
Dim arr_MSAT() As Long              '主分区表数组，指向的是存储分区表的SID
Dim arr_SAT() As Long               '分区表数组，指向的是下一个SID
Dim arr_SSAT() As Long              '短分区表数据
Dim arr_dir() As CFDir              '目录
'

Property Let path(file_path As String)
    my_path = file_path
End Property

Property Get path() As String
    path = my_path
End Property

Property Get ready() As Boolean
    ready = b_ready
End Property

Private Function getSSAT()
    Dim l_SID As Long
    Dim k As Long
   
    If cf_header.SSAT_count = 0 Then Exit Function
    '根目录的 stream_size 表示短流存放流的大小，每64个为一个short sector
    ReDim arr_SSAT(arr_dir(0).stream_size / 64 - 1) As Long
   
    l_SID = arr_dir(0).first_SID    '短流起始SID
    For k = 0 To arr_dir(0).stream_size / 64 - 1
        arr_SSAT(k) = l_SID * CFHEADER_SIZE + CFHEADER_SIZE + (k Mod 8) * 64  '指向偏移地址，实际地址要加上VarPtr(file_byte(0))
        If (k + 1) Mod 8 = 0 Then  '到下一个SID
            l_SID = arr_SAT(l_SID)
        End If
    Next k
   
    shtMSAT.Activate
    Range("H1").Resize(k, 1).Value = Application.WorksheetFunction.Transpose(arr_SSAT)

End Function

Private Function GetDir() As Long
    Dim l_sub_dir As Long
    Dim l_SID As Long
    Dim k As Long
   
    l_SID = cf_header.dir_first_SID
   
    k = 0
    Do
        ReDim Preserve arr_dir(k) As CFDir
        '首先是找到SID的地址，然后1个sector存放4个dir，1个偏移DIR_SIZE
        CopyMemory VarPtr(arr_dir(k).dir_name(0)), VarPtr(file_byte(0)) + CFHEADER_SIZE + CFHEADER_SIZE * l_SID + DIR_SIZE * (k Mod 4), DIR_SIZE
        k = k + 1
        If k Mod 4 = 0 Then
            l_SID = arr_SAT(l_SID)
        End If
    Loop Until l_SID = -2
   
End Function

Function DirToArr(arr() As String) As String
    ReDim arr(UBound(arr_dir), 13 - 1 - 1 - 1 + 1) As String '13个属性，-arr_keep,-notused +序号
    Dim i As Long
    Dim i_col As Long
    Dim k_dir As Long
   
    For k_dir = 0 To UBound(arr_dir)
        i_col = 0
        With arr_dir(k_dir)
            arr(k_dir, i_col) = my_hex(k_dir)
            i_col = i_col + 1
           
            arr(k_dir, i_col) = VBA.Left$(.dir_name, .len_name)
            i_col = i_col + 1
           
            arr(k_dir, i_col) = .len_name
            i_col = i_col + 1
           
            arr(k_dir, i_col) = .type
            i_col = i_col + 1
           
            arr(k_dir, i_col) = .color
            i_col = i_col + 1
                   
            arr(k_dir, i_col) = my_hex(.left_child)
            i_col = i_col + 1
           
            arr(k_dir, i_col) = my_hex(.right_child)
            i_col = i_col + 1
           
            arr(k_dir, i_col) = my_hex(.sub_dir)
            i_col = i_col + 1
                       
            arr(k_dir, i_col) = VBA.Format(.time_create, "yyyy/mm/dd")
            i_col = i_col + 1
           
            arr(k_dir, i_col) = VBA.Format(.time_modify, "yyyy/mm/dd")
            i_col = i_col + 1
           
            arr(k_dir, i_col) = my_hex(.first_SID)
            i_col = i_col + 1
           
            arr(k_dir, i_col) = my_hex(.stream_size)
            i_col = i_col + 1
        End With
    Next k_dir
   
End Function

Private Function GetSAT()
    Dim i As Long, j As Long
    Dim k_SAT As Long
    Dim arr(127) As Long
   
    With cf_header
        ReDim arr_SAT(.SAT_count * 128 - 1) As Long
       
        k_SAT = 0
        For i = 0 To .SAT_count - 1
            CopyMemory VarPtr(arr(0)), VarPtr(file_byte(0)) + CFHEADER_SIZE + CFHEADER_SIZE * arr_MSAT(i), CFHEADER_SIZE
           
            For j = 0 To 127
                arr_SAT(k_SAT) = arr(j)
                k_SAT = k_SAT + 1
            Next j
        Next i
       
        shtMSAT.Activate
        Range("E1").Resize(.SAT_count * 128, 1).Value = Application.WorksheetFunction.Transpose(arr_SAT)
   
    End With
End Function

Private Function GetMSAT()
    Dim i As Long
    Dim arr(127) As Long
    Dim next_SID As Long
    Dim flag As Boolean
    Dim count_MSAT As Long
   
    With cf_header
        ReDim arr_MSAT(.SAT_count - 1) As Long
       
        '获取头文件中的109个
        For i = 0 To 108
            If .arr_SID(i) = -1 Then
                '头中并没有109个
                Exit Function
            End If
           
            arr_MSAT(i) = .arr_SID(i)
        Next i
        '获取另外的
        count_MSAT = 109
        next_SID = .MSAT_first_SID
        flag = True
       
        Do
            CopyMemory VarPtr(arr(0)), VarPtr(file_byte(0)) + CFHEADER_SIZE + CFHEADER_SIZE * next_SID, CFHEADER_SIZE
           
            For i = 0 To 127 - 1
                If arr(i) = -1 Then
                    flag = False
                    Exit For
                End If
               
                arr_MSAT(count_MSAT) = arr(i)
                count_MSAT = count_MSAT + 1
            Next i
            next_SID = arr(i)       'SID的最后4个字节存储再下一个的SID
        Loop While flag
       
    End With
   
    shtMSAT.Activate
    Range("B1").Resize(count_MSAT, 1).Value = Application.WorksheetFunction.Transpose(arr_MSAT)
End Function

Private Function GetCfHeader()
    CopyMemory VarPtr(cf_header.id(0)), VarPtr(file_byte(0)), CFHEADER_SIZE
End Function

Function CfHeaderToArr(arr() As String) As String
    ReDim arr(1 + 16 + 108, 1) As String '1path,16个属性，108最后SID的下标
    Dim i As Long
    Dim i_row As Long
   
    i_row = 0
    With cf_header
        arr(i_row, 0) = "path"
        arr(i_row, 1) = Me.path
        i_row = i_row + 1
       
        arr(i_row, 0) = "cfID"
        For i = 0 To UBound(.id)
            arr(i_row, 1) = arr(i_row, 1) & my_hex(.id(i)) & " "
        Next i
        i_row = i_row + 1
       
        arr(i_row, 0) = "file_id"
        For i = 0 To UBound(.file_id)
            arr(i_row, 1) = arr(i_row, 1) & my_hex(.file_id(i)) & " "
        Next i
        i_row = i_row + 1
       
        arr(i_row, 0) = "file_format_revision"
        arr(i_row, 1) = .file_format_revision
        i_row = i_row + 1
       
        arr(i_row, 0) = "file_format_version"
        arr(i_row, 1) = .file_format_version
        i_row = i_row + 1
       
        arr(i_row, 0) = "memory_endian"
        arr(i_row, 1) = my_hex(.memory_endian)
        i_row = i_row + 1
       
        arr(i_row, 0) = "sector_size"
        arr(i_row, 1) = .sector_size
        i_row = i_row + 1
       
        arr(i_row, 0) = "short_sector_size"
        arr(i_row, 1) = .short_sector_size
        i_row = i_row + 1
       
        arr(i_row, 0) = "not_used_1"
        For i = 0 To UBound(.not_used_1)
            arr(i_row, 1) = arr(i_row, 1) & my_hex(.not_used_1(i)) & " "
        Next i
        i_row = i_row + 1
       
        arr(i_row, 0) = "SAT_count"
        arr(i_row, 1) = .SAT_count
        i_row = i_row + 1
                       
        arr(i_row, 0) = "dir_first_SID"
        arr(i_row, 1) = my_hex(.dir_first_SID)
        i_row = i_row + 1
       
        arr(i_row, 0) = "not_used_2"
        For i = 0 To UBound(.not_used_2)
            arr(i_row, 1) = arr(i_row, 1) & my_hex(.not_used_2(i)) & " "
        Next i
        i_row = i_row + 1
       
        arr(i_row, 0) = "min_stream_size"
        arr(i_row, 1) = my_hex(.min_stream_size)
        i_row = i_row + 1
               
        arr(i_row, 0) = "SSAT_first_SID"
        arr(i_row, 1) = my_hex(.SSAT_first_SID)
        i_row = i_row + 1
                               
        arr(i_row, 0) = "SSAT_count"
        arr(i_row, 1) = .SSAT_count
        i_row = i_row + 1
               
        arr(i_row, 0) = "MSAT_first_SID"
        arr(i_row, 1) = my_hex(.MSAT_first_SID)
        i_row = i_row + 1
               
        arr(i_row, 0) = "MSAT_count"
        arr(i_row, 1) = .MSAT_count
        i_row = i_row + 1
               
        arr(i_row, 0) = "arr_SID"
        For i = i_row To UBound(arr, 1)
            arr(i, 1) = my_hex(.arr_SID(i - i_row))
        Next i
               
    End With
   
End Function

Private Function GetFileByte() As Long
   
    Dim num_file As Integer
    Dim i_size As Long
   
    num_file = FreeFile
   
    Open Me.path For Binary Access Read As #num_file
    i_size = LOF(num_file)
    If i_size = 0 Then
        MsgBox "空文件。"
        GetFileByte = -1
        Exit Function
    End If
   
    ReDim file_byte(LOF(num_file) - 1) As Byte
    Get #num_file, 1, file_byte
    Close num_file
   
    GetFileByte = 1
End Function

Function read_stream(dir_name As String) As String
    Dim i As Long
    Dim str As String
    Dim arr_result() As Byte
    Dim n_size As Long, j As Long
    Dim short_SID As Long
   
    For i = 0 To UBound(arr_dir, 1)
        str = VBA.LeftB$(arr_dir(i).dir_name, arr_dir(i).len_name - 1)
        If VBA.StrComp(str, dir_name, vbTextCompare) = 0 Then Exit For
    Next i
   
    If i - 1 = UBound(arr_dir, 1) Then
        MsgBox "没有目录" & dir_name
        Exit Function
    End If
   
    If arr_dir(i).type <> 2 Then
        MsgBox "目录" & dir_name & "不是流。"
        Exit Function
    End If
   
    With arr_dir(i)
        If arr_dir(i).stream_size < cf_header.min_stream_size Then
            'short_sector
            n_size = .stream_size \ 64
           
            If .stream_size Mod 64 = 0 Then
                ReDim arr_result(.stream_size - 1) As Byte
            Else
                ReDim arr_result((n_size + 1) * 64 - 1) As Byte
            End If
           
            short_SID = .first_SID
            For j = 1 To n_size
'                Debug.Print VBA.Hex$(arr_SSAT(short_SID + j - 1))
                
                CopyMemory VarPtr(arr_result((j - 1) * 64)), VarPtr(file_byte(0)) + arr_SSAT(short_SID + j - 1), 64
            Next j
        Else
       
        End If
    End With
    
'    Dim arr() As String
'    ReDim arr(UBound(arr_result)) As String
'
'    Do While i <= UBound(arr_result)
'        If arr_result(i) < 128 Then
'            arr(i) = VBA.Chr(arr_result(i))
'            i = i + 1
'        Else
'            arr(i) = VBA.Chr("&H" & VBA.Hex$(arr_result(i)) & VBA.Hex$(arr_result(i + 1)))
'            i = i + 2
'        End If
'    Loop
    
'    read_stream = VBA.Join(arr, "") 'ByteToStr(arr_result, "GB2312") '
    
'    read_stream = ByteToStr(arr_result, "UTF-8")

    read_stream = VBA.StrConv(arr_result, vbUnicode, &H804)
End Function

Private Sub Class_Initialize()
    Me.path = GetFileName()
   
    If Me.path <> "" Then
        If GetFileByte = 1 Then
            b_ready = True
            GetCfHeader
           
            GetMSAT
           
            GetSAT
           
            GetDir
           
            getSSAT
           
        End If
    Else
        b_ready = False
    End If
End Sub

Private Sub Class_Terminate()
    Erase arr_MSAT
    Erase arr_SAT
    Erase arr_SSAT
    Erase file_byte
End Sub

