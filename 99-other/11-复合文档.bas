Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Type CompdocFileFormat '复合文档
    cff_identifier(7) As Byte '是复合文档的标识,依次是D0、CF、11、E0、A1、B1、1A、E1
    file_identifier(15) As Byte '文件唯一标识
    file_format_revision As Integer '文件格式的修订号
    file_format_version As Integer '文件格式的版本号
    memory_endian As Integer 'FFFEh表示“Little-Endian”
    sector_size As Integer '通常为9，扇区的大小,2的幂
    short_sector_size As Integer '通常为6，短扇区的大小，2的幂
    not_used_1(9) As Byte
    SAT_count As Long '分区表扇区的总数,“sector allocation table”（缩写SAT）
    first_sector_id As Long '目录流第一个扇区的ID
    not_used_2(3) As Byte '
    min_stream_size As Long '最小标准流尺寸，通常为1000h
    SSAT_first_id As Long '短分区表（SSAT）的第一个扇区的ID
    SSAT_count As Long '短分区表扇区总数
    MSAT_first_id As Long '主分区表（缩写MSAT）的第一个扇区的ID
    MSAT_count As Long '分区表的扇区总数

    '109个32位整数，为主分区表的开头109个记录
    '当分区表扇区SAT_count的个数大于109个时，就需要另外的扇区来存储，否则为-1
    arr_id(108) As Long
End Type

Type cff_dir
    dir_name(63) As Byte '目录入口名称
    name_len As Integer  '目录入口名称长度
    type As Byte        '1仓storage 2流 5根
    color As Byte       '0红色 1黑色
    left_child As Long  '-1表示叶子
    right_child As Long
    child_root As Long        '子目录红黑树的根节点，-1表示没有子目录
    arr_keep(19) As Byte    '权当保留
    dir_create As Date     '目录创建时间
    dir_modify As Date            '最后修改时间
    
    first_sid As Long   '目录入口所表示的流的第一个扇区编码
                        '对于根目录
                        
    stream_size As Long     '目录入口所表示的流的尺寸，通过将这个尺寸与短扇区尺寸进行比较，
                    '可以确定该流是是以扇区还是短扇区进行存储
    
    not_used As Long
End Type


Sub test()
    Dim cff As CompdocFileFormat
    Dim cff_dir As cff_dir
    Dim file_name As String
    Dim file_buffer() As Byte
    Dim arr_MSAT() As Long '主分区表数组对应每个分区表所在的 扇区ID
    Dim arr_SAT() As Long
    Dim arr_SSAT() As Long  'SSAT短分区表数组，存储扇区ID
    Dim l_SAT_count As Long '分区表个数
    Dim l_SID_count As Long '存储在SAT分区表中的扇区ID个数
    Dim l_SSID_count As Long '存储在SSAT短分区表中的扇区ID个数
    Dim i As Long
    
    file_name = ".xls"
    
    read_file file_name, file_buffer
    
    CopyMemory VarPtr(cff.cff_identifier(0)), VarPtr(file_buffer(0)), 512
    
    Dim temp As Long
    
    temp = cff.first_sector_id * 512 + 512
    Debug.Print VBA.Hex(temp)
    
    CopyMemory VarPtr(cff_dir.dir_name(0)), VarPtr(file_buffer(0)) + cff.first_sector_id * 512 + 512, 128 
    
    l_SAT_count = get_MSAT(cff, arr_MSAT) '获取主分区表数组,数组的值是 分区表SAT
    l_SID_count = get_SAT(file_buffer, cff, arr_MSAT, l_SAT_count, arr_SAT)
End Sub

'返回分区表SAT的个数
Function get_MSAT(cff As CompdocFileFormat, arr_MSAT() As Long) As Long
    Dim max_n_MSAT As Long '最多有多少个 分区表(主分区数组的个数)
    Dim i As Long
    
    max_n_MSAT = cff.MSAT_count * 127 + 109
    ReDim arr_MSAT(max_n_MSAT - 1) As Long
    
    i = 0
    '获取 头512个节点中的 109个记录
    Do While i < 108 And cff.arr_id(i) <> &HFFFFFFFF
        arr_MSAT(i) = cff.arr_id(i)
'        Debug.Print i, VBA.Hex(arr_MSAT(i))
        i = i + 1
    Loop
    
    If cff.SSAT_count > 0 Then
    
    End If
    
    get_MSAT = i
End Function

Function get_SAT(file_buffer() As Byte, cff As CompdocFileFormat, arr_MSAT() As Long, l_SAT_count As Long, arr_SAT() As Long) As Long
    Dim i As Long, j As Long
    Dim max_n_SAT As Long '最多有多少个 扇区
    Dim memory_offset As Long
    Dim arr_sector(512 / 4 - 1) As Long  '4个字节存1个 扇区ID
    Dim k_SID As Long                    '扇区个数
    
    max_n_SAT = l_SAT_count * 128
    ReDim arr_SAT(max_n_SAT - 1) As Long
    
    k_SID = 0
    For i = 0 To l_SAT_count - 1
        memory_offset = arr_MSAT(i) * 512 + 512
        CopyMemory VarPtr(arr_sector(0)), VarPtr(file_buffer(memory_offset)), 512
        For j = 0 To 127
'            If &HFFFFFFFE = arr_sector(j) Then Exit For '&HFFFFFFFE 扇区结束标志
            arr_SAT(k_SID) = arr_sector(j)
            k_SID = k_SID + 1
        Next j
    Next i
    
    get_SAT = k_SID
End Function

Function get_SSAT(file_buffer() As Byte, cff As CompdocFileFormat, arr_SSAT() As Long) As Long
    Dim i As Long, j As Long
    Dim max_n_SSAT As Long '最多有多少个 扇区
    Dim memory_offset As Long
    Dim arr_sector(512 / 4 - 1) As Long  '4个字节存1个 扇区ID
    Dim k_SID As Long                    '扇区个数
    
    max_n_SAT = cff.SSAT_count * 128
    ReDim arr_SSAT(max_n_SSAT - 1) As Long
    
    k_SID = 0
    For i = 0 To cff.SSAT_count - 1
        memory_offset = arr_MSAT(i) * 512 + 512
        CopyMemory VarPtr(arr_sector(0)), VarPtr(file_buffer(memory_offset)), 512
        For j = 0 To 127
'            If &HFFFFFFFE = arr_sector(j) Then Exit For '&HFFFFFFFE 扇区结束标志
            arr_SAT(k_SID) = arr_sector(j)
            k_SID = k_SID + 1
        Next j
    Next i
    
    get_SAT = k_SID
End Function

Function read_file(file_name As String, file_buffer() As Byte)
    Dim num_file As Integer
    
    num_file = FreeFile
    
    Open file_name For Binary Access Read As #num_file
    ReDim file_buffer(LOF(num_file) - 1) As Byte
    Get #num_file, 1, file_buffer
    Close num_file
End Function
