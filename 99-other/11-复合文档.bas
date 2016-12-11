Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Type CompdocFileFormat '�����ĵ�
    cff_identifier(7) As Byte '�Ǹ����ĵ��ı�ʶ,������D0��CF��11��E0��A1��B1��1A��E1
    file_identifier(15) As Byte '�ļ�Ψһ��ʶ
    file_format_revision As Integer '�ļ���ʽ���޶���
    file_format_version As Integer '�ļ���ʽ�İ汾��
    memory_endian As Integer 'FFFEh��ʾ��Little-Endian��
    sector_size As Integer 'ͨ��Ϊ9�������Ĵ�С,2����
    short_sector_size As Integer 'ͨ��Ϊ6���������Ĵ�С��2����
    not_used_1(9) As Byte
    SAT_count As Long '����������������,��sector allocation table������дSAT��
    first_sector_id As Long 'Ŀ¼����һ��������ID
    not_used_2(3) As Byte '
    min_stream_size As Long '��С��׼���ߴ磬ͨ��Ϊ1000h
    SSAT_first_id As Long '�̷�����SSAT���ĵ�һ��������ID
    SSAT_count As Long '�̷�������������
    MSAT_first_id As Long '����������дMSAT���ĵ�һ��������ID
    MSAT_count As Long '���������������

    '109��32λ������Ϊ��������Ŀ�ͷ109����¼
    '������������SAT_count�ĸ�������109��ʱ������Ҫ������������洢������Ϊ-1
    arr_id(108) As Long
End Type

Type cff_dir
    dir_name(63) As Byte 'Ŀ¼�������
    name_len As Integer  'Ŀ¼������Ƴ���
    type As Byte        '1��storage 2�� 5��
    color As Byte       '0��ɫ 1��ɫ
    left_child As Long  '-1��ʾҶ��
    right_child As Long
    child_root As Long        '��Ŀ¼������ĸ��ڵ㣬-1��ʾû����Ŀ¼
    arr_keep(19) As Byte    'Ȩ������
    dir_create As Date     'Ŀ¼����ʱ��
    dir_modify As Date            '����޸�ʱ��
    
    first_sid As Long   'Ŀ¼�������ʾ�����ĵ�һ����������
                        '���ڸ�Ŀ¼
                        
    stream_size As Long     'Ŀ¼�������ʾ�����ĳߴ磬ͨ��������ߴ���������ߴ���бȽϣ�
                    '����ȷ�������������������Ƕ��������д洢
    
    not_used As Long
End Type


Sub test()
    Dim cff As CompdocFileFormat
    Dim cff_dir As cff_dir
    Dim file_name As String
    Dim file_buffer() As Byte
    Dim arr_MSAT() As Long '�������������Ӧÿ�����������ڵ� ����ID
    Dim arr_SAT() As Long
    Dim arr_SSAT() As Long  'SSAT�̷��������飬�洢����ID
    Dim l_SAT_count As Long '���������
    Dim l_SID_count As Long '�洢��SAT�������е�����ID����
    Dim l_SSID_count As Long '�洢��SSAT�̷������е�����ID����
    Dim i As Long
    
    file_name = ".xls"
    
    read_file file_name, file_buffer
    
    CopyMemory VarPtr(cff.cff_identifier(0)), VarPtr(file_buffer(0)), 512
    
    Dim temp As Long
    
    temp = cff.first_sector_id * 512 + 512
    Debug.Print VBA.Hex(temp)
    
    CopyMemory VarPtr(cff_dir.dir_name(0)), VarPtr(file_buffer(0)) + cff.first_sector_id * 512 + 512, 128 
    
    l_SAT_count = get_MSAT(cff, arr_MSAT) '��ȡ������������,�����ֵ�� ������SAT
    l_SID_count = get_SAT(file_buffer, cff, arr_MSAT, l_SAT_count, arr_SAT)
End Sub

'���ط�����SAT�ĸ���
Function get_MSAT(cff As CompdocFileFormat, arr_MSAT() As Long) As Long
    Dim max_n_MSAT As Long '����ж��ٸ� ������(����������ĸ���)
    Dim i As Long
    
    max_n_MSAT = cff.MSAT_count * 127 + 109
    ReDim arr_MSAT(max_n_MSAT - 1) As Long
    
    i = 0
    '��ȡ ͷ512���ڵ��е� 109����¼
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
    Dim max_n_SAT As Long '����ж��ٸ� ����
    Dim memory_offset As Long
    Dim arr_sector(512 / 4 - 1) As Long  '4���ֽڴ�1�� ����ID
    Dim k_SID As Long                    '��������
    
    max_n_SAT = l_SAT_count * 128
    ReDim arr_SAT(max_n_SAT - 1) As Long
    
    k_SID = 0
    For i = 0 To l_SAT_count - 1
        memory_offset = arr_MSAT(i) * 512 + 512
        CopyMemory VarPtr(arr_sector(0)), VarPtr(file_buffer(memory_offset)), 512
        For j = 0 To 127
'            If &HFFFFFFFE = arr_sector(j) Then Exit For '&HFFFFFFFE ����������־
            arr_SAT(k_SID) = arr_sector(j)
            k_SID = k_SID + 1
        Next j
    Next i
    
    get_SAT = k_SID
End Function

Function get_SSAT(file_buffer() As Byte, cff As CompdocFileFormat, arr_SSAT() As Long) As Long
    Dim i As Long, j As Long
    Dim max_n_SSAT As Long '����ж��ٸ� ����
    Dim memory_offset As Long
    Dim arr_sector(512 / 4 - 1) As Long  '4���ֽڴ�1�� ����ID
    Dim k_SID As Long                    '��������
    
    max_n_SAT = cff.SSAT_count * 128
    ReDim arr_SSAT(max_n_SSAT - 1) As Long
    
    k_SID = 0
    For i = 0 To cff.SSAT_count - 1
        memory_offset = arr_MSAT(i) * 512 + 512
        CopyMemory VarPtr(arr_sector(0)), VarPtr(file_buffer(memory_offset)), 512
        For j = 0 To 127
'            If &HFFFFFFFE = arr_sector(j) Then Exit For '&HFFFFFFFE ����������־
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
