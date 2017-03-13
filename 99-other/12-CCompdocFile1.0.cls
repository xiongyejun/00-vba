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
    id(7) As Byte                   '�ĵ���ʶid
    file_id(15) As Byte             '�ļ�Ψһ��ʶ
    file_format_revision As Integer '�ļ���ʽ�޶���
    file_format_version As Integer  '�ļ���ʽ�汾��
    memory_endian As Integer        'FFFE��ʾ Little-Endian
    sector_size As Integer          '�����Ĵ�С 2���� ͨ��Ϊ2^9=512
    short_sector_size As Integer    '��������С��2����,ͨ��Ϊ2^6
    not_used_1(9) As Byte
    SAT_count As Long               '����������������
    dir_first_SID As Long           'Ŀ¼����һ��������ID
    not_used_2(3) As Byte
    min_stream_size As Long         '��С��׼��
    SSAT_first_SID As Long          '�̷������ĵ�һ������ID
    SSAT_count As Long              '�̷�������������
    MSAT_first_SID As Long          '���������ĵ�һ������ID
    MSAT_count As Long              '����������������
    arr_SID(108) As Long            '��������ǰ109����¼
End Type

Private Type CFDir
    dir_name(63) As Byte
    len_name As Integer
    type As Byte                    '1��storage 2�� 5��
    color As Byte                    '0��ɫ 1��ɫ
    left_child As Long               '-1��ʾҶ��
    right_child As Long
    sub_dir As Long
    arr_keep(19) As Byte
    time_create As Date
    time_modify As Date
    first_SID As Long               'Ŀ¼�������ʾ�ĵ�1����������
    stream_size As Long             'Ŀ¼������ߴ磬���ж��Ƿ��Ƕ�����
    not_used As Long
End Type

Private my_path As String
Private b_ready As Boolean
Dim file_byte() As Byte
Dim cf_header As CFHeader           '�ļ�ͷ
Dim arr_MSAT() As Long              '�����������飬ָ����Ǵ洢��������SID
Dim arr_SAT() As Long               '���������飬ָ�������һ��SID
Dim arr_SSAT() As Long              '�̷���������
Dim arr_dir() As CFDir              'Ŀ¼
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

Function write_file_byte(arr(), l_index As Long)
    Dim num_file As Integer
    Dim i As Long
    Dim k As Long
    
    num_file = FreeFile
    
    Open Me.path For Binary Access Write As #num_file
   
    k = 3
    
    Do Until arr(l_index, k) = 0
        If k Mod 2 = 1 Then
            Put #num_file, arr(l_index, k) + 1, VBA.CByte(&HD)
        Else
            Put #num_file, arr(l_index, k) + 1, VBA.CByte(&HA)
        End If
        k = k + 1
    Loop
    
    Close num_file
       
End Function

Function FindVba(arr()) As Long
    Dim str_PROJECT As String
    Dim re As Object
    Dim match_coll As Object
    Dim i As Long, j As Long, k As Long
    Dim arr_address() As Long
    Dim start_address As Long
    Dim l_double_byte As Long
    Dim this_double_byte As Long
    Dim if_short As Boolean
    Dim step_address As Long
    Dim l_max_len As Long
    
    str_PROJECT = Me.read_stream("PROJECT", arr_address, if_short)
    If if_short Then
        step_address = 64
    Else
        step_address = CFHEADER_SIZE
    End If
   
    Set re = CreateObject("VBScript.RegExp") 'Microsoft VBScript Tegular Expressions 5.5
    With re
        .Global = True                  '����ȫ���ַ���false��������1����ֹͣ
        .MultiLine = False              '�Ƿ����
        .IgnoreCase = False             '���ִ�Сд
        .Pattern = "Module=\S*|Class=\S*|BaseClass=\S*"      '���ع��� |Class=.{1,}||BaseClass=.{1,}|
        Set match_coll = .Execute(str_PROJECT)            '����MatchCollection����
    End With
    
    If match_coll.Count = 0 Then
        MsgBox "û�ҵ�ģ��"
        FindVba = 0
        Exit Function
    End If
    
    ReDim arr(match_coll.Count - 1, 2) '0-���ƣ�1-�ֽڿ�ʼ�ĵط���2-ռ�õ��ֽڸ�����3��nʵ��ÿ���ַ��ĵ�ַ
    l_double_byte = 0                       '˫�ֽ��ַ��ĸ���
    l_max_len = 0
    For i = 0 To match_coll.Count - 1
        arr(i, 0) = match_coll(i).Value     '����
        arr(i, 1) = match_coll(i).FirstIndex + l_double_byte '�ֽڿ�ʼ�ĵط�������˫�ֽ����
        this_double_byte = double_byte(VBA.CStr(arr(i, 0)))  '��ǰ�ַ���˫�ֽ��ַ�����
        l_double_byte = l_double_byte + this_double_byte
        arr(i, 2) = VBA.len(VBA.CStr(arr(i, 0))) + this_double_byte  'ռ�õ��ֽڸ���
        If arr(i, 2) > l_max_len Then l_max_len = arr(i, 2)
    Next i
    '������ַ����Ϊ�п����ǲ������ģ�������1��ģ�黹���ܿ��ܿ�Խ2��sector
    'ֱ�Ӽ��㵽ÿһ���ַ��ĵ�ַ
    ReDim Preserve arr(i - 1, 2 + l_max_len + 1) '+1Ϊ�˺�����ж�����λ��
    Dim p_address As Long                       '�����ĸ����ε��±���
    Dim byte_index As Long
    
    For j = 0 To i - 1
        byte_index = VBA.CDbl(arr(j, 1))
        
        For k = 0 To arr(j, 2) - 1
            p_address = (k + byte_index) \ step_address
            arr(j, k + 2 + 1) = arr_address(p_address, 1) + ((byte_index + k) Mod step_address)
        Next k
    Next j
    
    
   FindVba = i - 1
   
'    For j = 0 To i - 1
'        k = 3
'        Do Until arr(j, k) = 0
'            Debug.Print VBA.Hex$(arr(j, k))
'            k = k + 1
'        Loop
'    Next j
   
    Set re = Nothing
    Set match_coll = Nothing
End Function

Private Function getSSAT()
    Dim l_SID As Long
    Dim k As Long
   
    If cf_header.SSAT_count = 0 Then Exit Function
    '��Ŀ¼�� stream_size ��ʾ����������Ĵ�С��ÿ64��Ϊһ��short sector
    ReDim arr_SSAT(arr_dir(0).stream_size / 64 - 1) As Long
   
    l_SID = arr_dir(0).first_SID    '������ʼSID
    For k = 0 To arr_dir(0).stream_size / 64 - 1
        arr_SSAT(k) = l_SID * CFHEADER_SIZE + CFHEADER_SIZE + (k Mod 8) * 64  'ָ��ƫ�Ƶ�ַ��ʵ�ʵ�ַҪ����VarPtr(file_byte(0))
        If (k + 1) Mod 8 = 0 Then  '����һ��SID
            l_SID = arr_SAT(l_SID)
        End If
    Next k
   
'    shtMSAT.Activate
'    Range("H1").Resize(k, 1).Value = Application.WorksheetFunction.Transpose(arr_SSAT)

End Function

Private Function GetDir() As Long
    Dim l_sub_dir As Long
    Dim l_SID As Long
    Dim k As Long
   
    l_SID = cf_header.dir_first_SID
   
    k = 0
    Do
        ReDim Preserve arr_dir(k) As CFDir
        '�������ҵ�SID�ĵ�ַ��Ȼ��1��sector���4��dir��1��ƫ��DIR_SIZE
        CopyMemory VarPtr(arr_dir(k).dir_name(0)), VarPtr(file_byte(0)) + CFHEADER_SIZE + CFHEADER_SIZE * l_SID + DIR_SIZE * (k Mod 4), DIR_SIZE
        k = k + 1
        If k Mod 4 = 0 Then
            l_SID = arr_SAT(l_SID)
        End If
    Loop Until l_SID = -2
   
End Function

Function DirToArr(arr() As String) As String
    ReDim arr(UBound(arr_dir), 13 - 1 - 1 - 1 + 1) As String '13�����ԣ�-arr_keep,-notused +���
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
       
'        shtMSAT.Activate
'        Range("E1").Resize(.SAT_count * 128, 1).Value = Application.WorksheetFunction.Transpose(arr_SAT)
   
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
       
        '��ȡͷ�ļ��е�109��
        For i = 0 To 108
            If .arr_SID(i) = -1 Then
                'ͷ�в�û��109��
                Exit Function
            End If
           
            arr_MSAT(i) = .arr_SID(i)
        Next i
        '��ȡ�����
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
            next_SID = arr(i)       'SID�����4���ֽڴ洢����һ����SID
        Loop While flag
       
    End With
   
'    shtMSAT.Activate
'    Range("B1").Resize(count_MSAT, 1).Value = Application.WorksheetFunction.Transpose(arr_MSAT)
End Function

Private Function GetCfHeader()
    CopyMemory VarPtr(cf_header.id(0)), VarPtr(file_byte(0)), CFHEADER_SIZE
End Function

Function CfHeaderToArr(arr() As String) As String
    ReDim arr(1 + 16 + 108, 1) As String '1path,16�����ԣ�108���SID���±�
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
        MsgBox "���ļ���"
        GetFileByte = -1
        Exit Function
    End If
   
    ReDim file_byte(LOF(num_file) - 1) As Byte
    Get #num_file, 1, file_byte
    Close num_file
   
    GetFileByte = 1
End Function

'arr_address ����1�����ҵ�ַ�����飬�ڲ���ģ���ʱ���ã���Ϊ��������һ����������
'            ��1�д�����ǰstream�ֽ��±꣬��2���Ƕ�Ӧ�ĵ�ַ���磺
'           0   100
'           1   164
'           2   228
'Ҳ�п�����512��
Function read_stream(dir_name As String, arr_address() As Long, if_short As Boolean) As String
    Dim i As Long
    Dim str As String
    Dim arr_result() As Byte
    Dim n_size As Long, j As Long
    Dim short_SID As Long
    Dim l_SID As Long
    
    For i = 0 To UBound(arr_dir, 1)
        str = VBA.LeftB$(arr_dir(i).dir_name, arr_dir(i).len_name - 1)
        If VBA.StrComp(str, dir_name, vbTextCompare) = 0 Then Exit For
    Next i
   
    If i - 1 = UBound(arr_dir, 1) Then
        MsgBox "û��Ŀ¼" & dir_name
        Exit Function
    End If
   
    If arr_dir(i).type <> 2 Then
        MsgBox "Ŀ¼" & dir_name & "��������"
        Exit Function
    End If
    
     If arr_dir(i).first_SID = -1 Then
        MsgBox "Ŀ¼" & dir_name & "����СΪ0��"
        Exit Function
    End If
   
    With arr_dir(i)
        If arr_dir(i).stream_size < cf_header.min_stream_size Then
            'short_sector
            if_short = True
            n_size = .stream_size \ 64
            ReDim arr_address(n_size - 1, 1) As Long
            
            If .stream_size Mod 64 = 0 Then
                ReDim arr_result(.stream_size - 1) As Byte
            Else
                ReDim arr_result((n_size + 1) * 64 - 1) As Byte
            End If
           
            short_SID = .first_SID
'            start_address = VarPtr(file_byte(0)) + arr_SSAT(short_SID)
            
            For j = 1 To n_size
                arr_address(j - 1, 0) = (j - 1)
                arr_address(j - 1, 1) = arr_SSAT(short_SID + j - 1) 'VarPtr(file_byte(0))
                CopyMemory VarPtr(arr_result((j - 1) * 64)), arr_address(j - 1, 1) + VarPtr(file_byte(0)), 64

            Next j
        Else
            if_short = False
            n_size = .stream_size \ CFHEADER_SIZE '512
            ReDim arr_address(n_size - 1, 1) As Long
            
            If .stream_size Mod CFHEADER_SIZE = 0 Then
                ReDim arr_result(.stream_size - 1) As Byte
            Else
                ReDim arr_result((n_size + 1) * CFHEADER_SIZE - 1) As Byte
            End If
            
            l_SID = .first_SID
'            start_address = VarPtr(file_byte(0)) + arr_SAT(l_SID) * CFHEADER_SIZE + CFHEADER_SIZE
            For j = 1 To n_size
                arr_address(j - 1, 0) = j - 1
                arr_address(j - 1, 1) = arr_SAT(l_SID) * CFHEADER_SIZE + CFHEADER_SIZE  '+VarPtr(file_byte(0))
                CopyMemory VarPtr(arr_result((j - 1) * CFHEADER_SIZE)), arr_address(j - 1, 1) + VarPtr(file_byte(0)), CFHEADER_SIZE
            Next j
            
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

Function my_hex(num As Variant) As String
    
    my_hex = VBA.Hex$(num)
    
    If VBA.len(my_hex) = 1 Then
        my_hex = "0" & my_hex
    End If
    
    my_hex = "&H" & my_hex
End Function

Function double_byte(str As String) As Long
    Dim i As Long
    For i = 1 To VBA.len(str)
        If VBA.Asc(VBA.Mid$(str, i, 1)) < 0 Then
            double_byte = double_byte + 1
        End If
    Next i
End Function