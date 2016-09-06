[http://club.excelhome.net/forum.php?mod=viewthread&tid=629386&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=629386&extra=page%3D1)

大家都知道，VBA中没有像VB那样提供Timer控件，而是用一个简单的onTime方法来实现定时功能。论坛里也有一些使用API函数构造的Timer类，但是因为AddressOf函数的问题，有些Timer类只能实现一个实例运行。

Joforn之前在趣味作品大赛里分享了一个扫雷游戏。里面使用了大量的API函数构造一个完美的Windows窗体程序，其中也使用了Paul Caton的类模块Call Back函数。有了这个类模块CallBack函数，便可以在类模块中实现AddressOf函数的功能。

根据这个方法，做了一个Timer类，实现了和VB中Timer控件完全一样的功能。
属性：

- Enabled：Boolean类型；设为True，启动计时器；设为False，则关闭计时器，默认为False。
- Interval：Long类型；计时器间隔时间，单位为毫秒，默认为0。
- 事件：Timer

有兴趣的试试看。注意在关闭文件之前需要停止Timer类，否则会导致Excel崩溃。

	'*************************************************************************************************
	'* clsTimer 1.0 - Timer class module
	'* ----------------------------------
	'*
	'*
	'*
	'* Use Paul Caton's cCallBack class module
	'*************************************************************************************************
	Option Explicit
	
	' API function of Timer process
	Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, _
	                        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
	Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
	
	Private bEnable As Boolean
	Private lDuration As Long
	Private lTimerId As Long
	Private lTimerProc As Long
	
	' Event
	Public Event Timer()
	
	'-Callback declarations for Paul Caton thunking magic----------------------------------------------
	Private z_CbMem   As Long    'Callback allocated memory address
	Private z_Cb()    As Long    'Callback thunk array
	
	Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
	Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
	Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
	Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
	Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
	Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
	'-------------------------------------------------------------------------------------------------
	
	Public Property Let Enabled(ByVal vData As Boolean)
	    bEnable = vData
	    If bEnable = True Then
	        StartTimer
	    Else
	        EndTimer
	    End If
	End Property
	
	Public Property Get Enabled() As Boolean
	    Enabled = bEnable
	End Property
	
	Public Property Let Interval(ByVal vData As Long)
	    If vData < 0 Then vData = 0
	    lDuration = vData
	    If lDuration > 0 And bEnable = True Then    ' If change the interval, stop the timer first, and start again
	        EndTimer
	        StartTimer
	    End If
	End Property
	
	Public Property Get Interval() As Long
	    Interval = lDuration
	End Property
	
	Private Sub Class_Initialize()
	    bEnable = False
	    lDuration = 0
	    lTimerId = 0
	    lTimerProc = 0
	End Sub
	
	Private Sub Class_Terminate()
	    bEnable = False
	    lDuration = 0
	    lTimerId = 0
	    lTimerProc = 0
	    zTerminate
	End Sub
	
	Private Sub StartTimer()
	    If lTimerProc = 0 And bEnable = True And lDuration > 0 Then
	        ' get address of timer process
	        lTimerProc = zb_AddressOf(1, 4)
	        ' start timer, return timer ID
	        lTimerId = SetTimer(0&, 0&, lDuration, lTimerProc)
	    End If
	End Sub
	
	Private Sub EndTimer()
	    If lTimerProc Then
	        KillTimer 0&, lTimerId
	        lTimerId = 0
	        lTimerProc = 0
	    End If
	End Sub
	
	'-------------------------------------------------------------------------------------------------
	'*************************************************************************************************
	'* cCallback - Class generic callback template
	'*
	'* Note:
	'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
	'*  The callback declarations and code can co-exist with subclassing declarations and code.
	'*    With both types of code in a single file,..
	'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
	'*      pay careful attention to the nOrdinal parameter to zAddressOf
	'*
	'* Paul_Caton@hotmail.com
	'* Copyright free, use and abuse as you see fit.
	'*
	'* v1.0 The original..................................................................... 20060408
	'* v1.1 Added multi-thunk support........................................................ 20060409
	'* v1.2 Added optional IDE protection.................................................... 20060411
	'* v1.3 Added an optional callback target object......................................... 20060413
	'*************************************************************************************************
	
	'-Callback code-----------------------------------------------------------------------------------
	Private Function zb_AddressOf(ByVal nOrdinal As Long, _
	                              ByVal nParamCount As Long, _
	                     Optional ByVal nThunkNo As Long = 0, _
	                     Optional ByVal oCallback As Object = Nothing, _
	                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
	'*************************************************************************************************
	'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
	'* nParamCount  - The number of parameters that will callback
	'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
	'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
	'* bIdeSafety   - Optional, set to false to disable IDE protection.
	'*************************************************************************************************
	Const MAX_FUNKS   As Long = 1                                               'Number of simultaneous thunks, adjust to taste
	Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
	Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
	Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
	Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
	Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
	  Dim nAddr       As Long
	  
	  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
	    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
	    Exit Function
	  End If
	  
	  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
	    Set oCallback = Me                                                      'Then it is me
	  End If
	  
	  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
	  If nAddr = 0 Then
	    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
	    Exit Function
	  End If
	  
	  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
	    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
	    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
	  End If
	  
	  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
	    z_Cb(3, nThunkNo) = _
	              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
	    z_Cb(4, nThunkNo) = &HBB60E089
	    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
	    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
	  End If
	  
	  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
	  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
	  
	  If bIdeSafety Then                                                        'If the user wants IDE protection
	    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
	  End If
	    
	  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
	  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
	  
	  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
	  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
	  zb_AddressOf = nAddr + 16                                                 'Thunk code start address
	End Function
	
	'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
	Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
	  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
	  Dim bVal  As Byte
	  Dim nAddr As Long                                                         'Address of the vTable
	  Dim I     As Long                                                         'Loop index
	  Dim J     As Long                                                         'Loop limit
	  
	  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
	  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
	    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
	      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
	        Exit Function                                                       'Bail...
	      End If
	    End If
	  End If
	  
	  I = I + 4                                                                 'Bump to the next entry
	  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
	  Do While I < J
	    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
	    
	    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
	      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
	      Exit Do                                                               'Bad method signature, quit loop
	    End If
	
	    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
	    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
	      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
	      Exit Do                                                               'Bad method signature, quit loop
	    End If
	    
	    I = I + 4                                                             'Next vTable entry
	  Loop
	End Function
	
	'Probe at the specified start address for a method signature
	Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
	  Dim bVal    As Byte
	  Dim nAddr   As Long
	  Dim nLimit  As Long
	  Dim nEntry  As Long
	  
	  nAddr = nStart                                                            'Start address
	  nLimit = nAddr + 32                                                       'Probe eight entries
	  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
	    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
	    
	    If nEntry <> 0 Then                                                     'If not an implemented interface
	      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
	      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
	        nMethod = nAddr                                                     'Store the vTable entry
	        bSub = bVal                                                         'Store the found method signature
	        zProbe = True                                                       'Indicate success
	        Exit Function                                                       'Return
	      End If
	    End If
	    
	    nAddr = nAddr + 4                                                       'Next vTable entry
	  Loop
	End Function
	
	Private Sub zTerminate()
	    
	    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
	    If Not z_CbMem = 0 Then                                            'If memory allocated
	        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
	            z_CbMem = 0  'Release; Indicate memory released
	            Erase z_Cb()
	        End If
	    End If
	End Sub
	
	'*************************************************************************************************
	'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
	'*************************************************************************************************
	
	'Callback ordinal 2
	'Private Function NewWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	'
	'
	'End Function
	
	'Callback ordinal 1
	Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
	    RaiseEvent Timer
	End Function
	
工作表代码

	Private WithEvents cTimer1 As clsTimer
	Private WithEvents cTimer2 As clsTimer
	
	Private Sub CommandButton1_Click()
	    Set cTimer1 = New clsTimer
	    cTimer1.Interval = 100
	    cTimer1.Enabled = True
	    Set cTimer2 = New clsTimer
	    cTimer2.Interval = 200
	    cTimer2.Enabled = True
	End Sub
	
	Private Sub CommandButton2_Click()
	    cTimer1.Enabled = False
	    Set cTimer1 = Nothing
	End Sub
	
	Private Sub CommandButton3_Click()
	    cTimer2.Enabled = False
	    Set cTimer2 = Nothing
	End Sub
	
	Private Sub CommandButton4_Click()
	    cTimer2.Interval = 4000
	End Sub
	
	Private Sub cTimer1_Timer()
	    Cells(1, 1).Value = Cells(1, 1).Value + 1
	End Sub
	
	Private Sub cTimer2_Timer()
	    Cells(2, 1).Value = Cells(2, 1).Value + 1
	End Sub
