VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTrickTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' //
' // 64/32 bit timer class for VBA
' // by The trick 2019
' //

Option Explicit
Option Base 0

Private Const FADF_AUTO                     As Long = 1
Private Const HEAP_CREATE_ENABLE_EXECUTE    As Long = &H40000
Private Const HEAP_NO_SERIALIZE             As Long = &H1
Private Const HEAP_ZERO_MEMORY              As Long = &H8
Private Const PROCESS_HEAP_ENTRY_BUSY       As Long = &H4
Private Const WNDPROCINDEX                  As Long = 8
Private Const HEAP_ENV_VARIABLE             As String = "TrickVBATimer"
Private Const TIMERPROC_INDEX               As Long = 5

Private Type SAFEARRAYBOUND
    cElements           As Long
    lLbound             As Long
End Type

#If VBA7 Then

Private Type PROCESS_HEAP_ENTRY
    lpData              As LongPtr
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As LongPtr
    lpLastBlock         As LongPtr
End Type

Private Type SAFEARRAY
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    Bounds              As SAFEARRAYBOUND
End Type

Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" _
                                 Alias "SetEnvironmentVariableW" ( _
                                 ByVal lpName As LongPtr, _
                                 ByVal lpValue As LongPtr) As Long
Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" _
                                 Alias "GetEnvironmentVariableW" ( _
                                 ByVal lpName As LongPtr, _
                                 ByVal lpBuffer As LongPtr, _
                                 ByVal nSize As Long) As Long
Private Declare PtrSafe Function HeapCreate Lib "kernel32" ( _
                                 ByVal flOptions As Long, _
                                 ByVal dwInitialSize As LongPtr, _
                                 ByVal dwMaximumSize As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapDestroy Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr) As Long
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr, _
                                 ByVal dwFlags As Long, _
                                 ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr, _
                                 ByVal dwFlags As Long, _
                                 ByVal lpMem As LongPtr) As Long
Private Declare PtrSafe Function HeapWalk Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr, _
                                 ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare PtrSafe Function HeapLock Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr) As Long
Private Declare PtrSafe Function HeapUnlock Lib "kernel32" ( _
                                 ByVal hHeap As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib _
                                 "kernel32" Alias "GetModuleHandleW" ( _
                                 ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
                                 ByVal hModule As LongPtr, _
                                 ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" _
                                 Alias "LoadLibraryW" ( _
                                 ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
                                 ByVal hwnd As LongPtr, _
                                 ByVal nIDEvent As LongPtr, _
                                 ByVal uElapse As Long, _
                                 ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
                                 ByVal hwnd As LongPtr, _
                                 ByVal nIDEvent As LongPtr) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
                            Alias "RtlMoveMemory" ( _
                            ByRef pDestination As Any, _
                            ByRef Source As Any, _
                            ByVal Length As LongPtr)
#If Win64 Then
Private Declare PtrSafe Sub DupArray Lib "kernel32" _
                            Alias "RtlMoveMemory" ( _
                            ByRef Destination() As Any, _
                            ByRef pSA As Any, _
                            Optional ByVal Length As LongPtr = 8)
#Else
Private Declare PtrSafe Sub DupArray Lib "kernel32" _
                            Alias "RtlMoveMemory" ( _
                            ByRef Destination() As Any, _
                            ByRef pSA As Any, _
                            Optional ByVal Length As LongPtr = 4)
#End If

Private m_pAsmThunk     As LongPtr
Private m_hCodeHeap     As LongPtr
Private m_pEbMode       As LongPtr
Private m_lIdEvent      As LongPtr

#Else

Private Type PROCESS_HEAP_ENTRY
    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long
End Type

Private Type SAFEARRAY
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    Bounds              As SAFEARRAYBOUND
End Type

Private Declare Function SetEnvironmentVariable Lib "kernel32" _
                         Alias "SetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpValue As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" _
                         Alias "GetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpBuffer As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function HeapCreate Lib "kernel32" ( _
                         ByVal flOptions As Long, _
                         ByVal dwInitialSize As Long, _
                         ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMem As Long) As Long
Private Declare Function HeapWalk Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare Function HeapLock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function SetTimer Lib "user32" ( _
                         ByVal hwnd As Long, _
                         ByVal nIDEvent As Long, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" ( _
                         ByVal hwnd As Long, _
                         ByVal nIDEvent As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef pDestination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub DupArray Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination() As Any, _
                    ByRef pSA As Any, _
                    Optional ByVal Length As Long = 4)
                    
Private m_pAsmThunk     As Long
Private m_hCodeHeap     As Long
Private m_pEbMode       As Long
Private m_lIdEvent      As Long

#End If

Public Event Tick()

Private m_vTag          As Variant
Private m_lInterval     As Long

Public Property Let Interval( _
                    ByVal lValue As Long)
    
    If lValue = m_lInterval Then Exit Property
    
    If m_pAsmThunk Then
        
        KillTimer 0, m_lIdEvent
        
        m_lIdEvent = SetTimer(0, 0, lValue, m_pAsmThunk)
        If m_lIdEvent = 0 Then
            Err.Raise 5
        End If

#If Win64 Then
        CopyMemory ByVal m_pAsmThunk + &H3A, m_lIdEvent, Len(m_lIdEvent)
#Else
        CopyMemory ByVal m_pAsmThunk + &H16, m_lIdEvent, Len(m_lIdEvent)
#End If

    Else
        If Not CreateAsm(lValue) Then
            Err.Raise 5
        End If
    End If
    
    m_lInterval = lValue
                        
End Property

Public Property Get Interval() As Long
    Interval = m_lInterval
End Property

Public Property Let Tag( _
                    ByVal vValue As Variant)
    m_vTag = vValue
End Property

Public Property Set Tag( _
                    ByVal vValue As Variant)
    Set m_vTag = vValue
End Property

Public Property Get Tag() As Variant
    
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If
    
End Property

' // Callback function
' // If you modify position of this function you should update TIMERPROC_INDEX constant
#If VBA7 Then
Private Function TimerProc( _
                 ByVal hwnd As LongPtr, _
                 ByVal uMsg As Long, _
                 ByVal idEvent As LongPtr, _
                 ByVal dwTime As Long) As Long
#Else
Private Function TimerProc( _
                 ByVal hwnd As Long, _
                 ByVal uMsg As Long, _
                 ByVal idEvent As Long, _
                 ByVal dwTime As Long) As Long
#End If
    On Error GoTo error_handler
    
    RaiseEvent Tick
    
error_handler:
    
End Function
    
' // Create assembly thunk if doesn't exists
Private Function CreateAsm( _
                 ByVal lInterval As Long) As Boolean
    Dim bIsInIDE    As Boolean
#If VBA7 Then
    Dim lIdEvent    As LongPtr
#Else
    Dim lIdEvent    As Long
#If Not CBool(VBA6) Then
    
    Debug.Assert MakeTrue(bIsInIDE)
    
#End If
#End If

    If m_pAsmThunk Then
        CreateAsm = True
        Exit Function
    End If
    
    If GetCodeHeap() = 0 Then Exit Function

#If Not CBool(VBA6 Or VBA7) Then
    If bIsInIDE Then
#End If

    If m_pEbMode = 0 Then
        m_pEbMode = SearchEbMode()
        If m_pEbMode = 0 Then Exit Function
    End If
    
#If Not CBool(VBA6 Or VBA7) Then
    End If
#End If

#If VBA7 And Win64 Then
    
    m_pAsmThunk = Create64BitThunk()
    If m_pAsmThunk = 0 Then Exit Function
    
    lIdEvent = SetTimer(0, 0, lInterval, m_pAsmThunk)
    If lIdEvent = 0 Then
        
        HeapFree m_hCodeHeap, HEAP_NO_SERIALIZE, m_pAsmThunk - 4
        m_pAsmThunk = 0
        Exit Function
    
    End If
    
    CopyMemory ByVal m_pAsmThunk + &H3A, lIdEvent, Len(lIdEvent)
    
#Else

    m_pAsmThunk = Create32BitThunk()
    If m_pAsmThunk = 0 Then Exit Function
    
    lIdEvent = SetTimer(0, 0, lInterval, m_pAsmThunk)
    If lIdEvent = 0 Then
        
        HeapFree m_hCodeHeap, HEAP_NO_SERIALIZE, m_pAsmThunk - 4
        m_pAsmThunk = 0
        Exit Function
    
    End If
#If VBA6 Or VBA7 Then
    CopyMemory ByVal m_pAsmThunk + &H16, lIdEvent, Len(lIdEvent)
#Else
    If bIsInIDE Then
        CopyMemory ByVal m_pAsmThunk + &H16, lIdEvent, Len(lIdEvent)
    Else
    End If
#End If

    ' // Debug.Print Hex(m_pAsmThunk)
    
#End If
    
    m_lIdEvent = lIdEvent
    
    CreateAsm = True
    
End Function

#If VBA7 And Win64 Then
' // Search for EbMode function
Private Function SearchEbMode() As LongPtr
    Dim hVbe            As LongPtr
    Dim e_lfanew        As Long
    Dim iNumOfSec       As Integer
    Dim iOptSize        As Integer
    Dim pSection        As LongPtr
    Dim lIndex          As Long
    Dim cName           As Currency
    Dim pStartScan      As LongPtr
    Dim pEndScan        As LongPtr
    Dim bTemplate(&H5F) As Byte
    Dim bMask(&H5F)     As Byte
    Dim bData()         As Byte
    Dim tSAMap          As SAFEARRAY
    
    ' // 0. Try to get EbMode from previous thunks
    SearchEbMode = GetEbModeFromThunks()
    
    If SearchEbMode Then
        Exit Function
    End If
    
    ' // 1. Search for VBE7.dll ".text" section
    
    hVbe = GetModuleHandle(StrPtr("VBE7"))
    If hVbe = 0 Then Exit Function
    
    CopyMemory e_lfanew, ByVal hVbe + &H3C, 4
    CopyMemory iNumOfSec, ByVal hVbe + e_lfanew + 6, 2
    CopyMemory iOptSize, ByVal hVbe + e_lfanew + &H14, 2
    
    pSection = hVbe + e_lfanew + &H18 + iOptSize
    
    For lIndex = 0 To iNumOfSec - 1
        
        CopyMemory cName, ByVal pSection, 8
        
        ' // Search for ".text" section
        If cName = 50023612.1134@ Then
            
            CopyMemory pStartScan, ByVal pSection + &HC, 4
            CopyMemory pEndScan, ByVal pSection + &H8, 4
            
            pStartScan = pStartScan + hVbe
            pEndScan = pEndScan + pStartScan - 1
            
            Exit For
            
        End If
        
        pSection = pSection + &H28
        
    Next
    
    If pStartScan = 0 Or pEndScan = 0 Then Exit Function
    
    ' // 2. Search for Proc/MethCallEngine thunk template
    
    ' // 48 89 4C 24 08 48 89 54 24 10 4C 89 44 24 18 4C
    ' // 89 4C 24 20 48 B8 11 11 11 11 11 11 11 11 48 0B
    ' // C0 74 32 48 B8 XX XX XX XX XX XX XX XX FF D0 48
    ' // 83 F8 02 74 20 48 B8 11 11 11 11 11 11 11 11 48
    ' // 8B 4C 24 08 48 8B 54 24 10 4C 8B 44 24 18 4C 8B
    ' // 4C 24 20 FF E0 48 33 C0 C2 11 11
    
    ' // XX XX XX XX XX XX XX XX - EbMode
    
    ' // Setup template
    CopyMemory bTemplate(&H0), 609147917080124.0392@, 8
    CopyMemory bTemplate(&H8), 548317242310341.8404@, 8
    CopyMemory bTemplate(&H10), 122996679316526.1961@, 8
    CopyMemory bTemplate(&H18), 81291849773882.1905@, 8
    CopyMemory bTemplate(&H20), 79148524.8704@, 8
    CopyMemory bTemplate(&H28), 524697394135171.072@, 8
    CopyMemory bTemplate(&H30), 127684979858204.0707@, 8
    CopyMemory bTemplate(&H38), 519295061033333.9921@, 8
    CopyMemory bTemplate(&H40), 261787042489960.3595@, 8
    CopyMemory bTemplate(&H48), -840931986015968.9712@, 8
    CopyMemory bTemplate(&H50), -459725066342497.3748@, 8
    CopyMemory bTemplate(&H58), 111.8658@, 8
    
    ' // Setup mask
    For lIndex = 0 To UBound(bMask)
        If lIndex < &H25 Or (lIndex > &H2C And lIndex < &H5B) Then
            bMask(lIndex) = 1
        End If
    Next
    
    ' // Map array to data
    tSAMap.cbElements = 1
    tSAMap.cDims = 1
    tSAMap.fFeatures = FADF_AUTO
    tSAMap.Bounds.cElements = CLng(pEndScan - pStartScan) + 1
    tSAMap.pvData = pStartScan

    DupArray bData, VarPtr(tSAMap)
    
    lIndex = FindSignature(bData(), bTemplate(), bMask())
    
    DupArray bData, 0@
    
    If lIndex = -1 Then Exit Function
    
    CopyMemory SearchEbMode, ByVal pStartScan + lIndex + &H25, Len(SearchEbMode)
    
End Function

' // Create 64 bit thunk
Private Function Create64BitThunk() As LongPtr
    Dim pCode           As LongPtr
    Dim llThunk(19)     As Currency
    Dim pfnKillTimer    As LongPtr
    Dim pfnTimerProc    As LongPtr
    Dim pVtbl           As LongPtr
    Dim hUser32         As LongPtr
    
    If m_hCodeHeap = 0 Then Exit Function
    
    hUser32 = GetModuleHandle(StrPtr("user32"))
    If hUser32 = 0 Then
    
        hUser32 = LoadLibrary(StrPtr("user32"))
        If hUser32 = 0 Then
            Exit Function
        End If
        
    End If
    
    pfnKillTimer = GetProcAddress(hUser32, "KillTimer")
    If pfnKillTimer = 0 Then Exit Function
    
    pCode = HeapAlloc(m_hCodeHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, &H9E)
    If pCode = 0 Then Exit Function
    
    pCode = pCode + 4   ' // Disable
    
    ' // 48 83 EC 38 48 89 4C 24 40 48 89 54 24 48 4C 89
    ' // 44 24 50 4C 89 4C 24 58 48 B8 00 00 00 00 00 00
    ' // 00 00 FF D0 48 85 C0 74 06 3C 01 74 23 EB 66 FF
    ' // 0D C7 FF FF FF 48 31 C9 48 BA 00 00 00 00 00 00
    ' // 00 00 48 B8 00 00 00 00 00 00 00 00 FF D0 EB 45
    ' // 48 B9 00 00 00 00 00 00 00 00 48 8B 54 24 40 4C
    ' // 8B 44 24 48 4C 8B 4C 24 50 48 8B 44 24 58 48 89
    ' // 44 24 20 48 8D 44 24 30 48 C7 00 00 00 00 00 48
    ' // 89 44 24 28 48 B8 00 00 00 00 00 00 00 00 FF D0
    ' // 48 8B 44 24 30 48 83 C4 38 C3

    llThunk(0) = 261561642688109.0376@
    llThunk(1) = -855338227140910.8928@
    llThunk(2) = 635128552707379.3092@
    llThunk(3) = 4.7176@
    llThunk(4) = 841287065171859.8656@
    llThunk(5) = -4308860891082.0346@
    llThunk(6) = -394929513387840.1267@
    llThunk(7) = 4.7688@
    llThunk(8) = 309172.6336@
    llThunk(9) = 503835040177926.9632@
    llThunk(10) = 4.7432@
    llThunk(11) = 549443149092460.9536@
    llThunk(12) = 261561864333952.7307@
    llThunk(13) = -855449057939821.1504@
    llThunk(14) = 346897298657326.8036@
    llThunk(15) = 518814677073086.2408@
    llThunk(16) = 20262005062.1577@
    llThunk(17) = -338698839475932.3648@
    llThunk(18) = -428650304872247.8264@
    llThunk(19) = 4.9976@
    
    ' // Get TimerProc address
    CopyMemory pVtbl, ByVal ObjPtr(Me), Len(pVtbl)

    CopyMemory pfnTimerProc, ByVal pVtbl + (TIMERPROC_INDEX + 7) * Len(pfnTimerProc), Len(pfnTimerProc)
    
    CopyMemory ByVal pCode, llThunk(0), &H9A
    CopyMemory ByVal pCode + &H1A, m_pEbMode, Len(m_pEbMode)
    CopyMemory ByVal pCode + &H44, pfnKillTimer, Len(pfnKillTimer)
    CopyMemory ByVal pCode + &H52, ObjPtr(Me), 8
    CopyMemory ByVal pCode + &H86, pfnTimerProc, Len(pfnTimerProc)
    
    Create64BitThunk = pCode
    
End Function

#Else

' // Search for EbMode function
#If VBA7 Then
Private Function SearchEbMode() As LongPtr
    Dim hVbe            As LongPtr
    Dim pSection        As LongPtr
    Dim pStartScan      As LongPtr
    Dim pEndScan        As LongPtr
#Else
Private Function SearchEbMode() As Long
    Dim hVbe            As Long
    Dim pSection        As Long
    Dim pStartScan      As Long
    Dim pEndScan        As Long
#End If
    Dim e_lfanew        As Long
    Dim iNumOfSec       As Integer
    Dim iOptSize        As Integer
    Dim lIndex          As Long
    Dim cName           As Currency
    Dim bTemplate(&H27) As Byte
    Dim bMask(&H27)     As Byte
    Dim bData()         As Byte
    Dim tSAMap          As SAFEARRAY

    ' // 0. Try to get EbMode from previous thunks
    SearchEbMode = GetEbModeFromThunks()
    
    If SearchEbMode Then
        Exit Function
    End If
    
    ' // 1. Search for VBE7/6.dll ".text" section
    
#If VBA7 Then
    hVbe = GetModuleHandle(StrPtr("VBE7"))
#ElseIf VBA6 Then
    hVbe = GetModuleHandle(StrPtr("VBE6"))
#Else
    hVbe = GetModuleHandle(StrPtr("VBA6"))
#End If

    If hVbe = 0 Then Exit Function
    
#If VBA6 Or VBA7 Then

    CopyMemory e_lfanew, ByVal hVbe + &H3C, 4
    CopyMemory iNumOfSec, ByVal hVbe + e_lfanew + 6, 2
    CopyMemory iOptSize, ByVal hVbe + e_lfanew + &H14, 2
    
    pSection = hVbe + e_lfanew + &H18 + iOptSize
    
    For lIndex = 0 To iNumOfSec - 1
        
        CopyMemory cName, ByVal pSection, 8
        
        ' // Search for ".text" section
        If cName = 50023612.1134@ Then
            
            CopyMemory pStartScan, ByVal pSection + &HC, 4
            CopyMemory pEndScan, ByVal pSection + &H8, 4
            
            pStartScan = pStartScan + hVbe
            pEndScan = pEndScan + pStartScan - 1
            
            Exit For
            
        End If
        
        pSection = pSection + &H28
        
    Next
    
    If pStartScan = 0 Or pEndScan = 0 Then Exit Function
    
    ' // 2. Search for Proc/MethCallEngine thunk template
    
    ' // A1 YY YY YY YY 0B C0 74 13 B8 XX XX XX XX FF D0
    ' // 83 F8 02 74 07 B8 11 11 11 11 FF E0 33 C0 C2 11
    ' // 11
    
    ' // XX XX XX XX - EbMode
    
    ' // Setup template
    CopyMemory bTemplate(&H0), 841273619855599.2225@, 8
    CopyMemory bTemplate(&H8), -338698839475927.6525@, 8
    CopyMemory bTemplate(&H10), 122996651539948.9667@, 8
    CopyMemory bTemplate(&H18), 127979657317731.9697@, 8
    CopyMemory bTemplate(&H20), 0.0017@, 8

    ' // Setup mask
    For lIndex = 0 To UBound(bMask)
    
        Select Case lIndex
        Case 1 To 4, 10 To 13, 33 To 40
            bMask(lIndex) = 0
        Case Else
            bMask(lIndex) = 1
        End Select

    Next
    
    ' // Map array to data
    tSAMap.cbElements = 1
    tSAMap.cDims = 1
    tSAMap.fFeatures = FADF_AUTO
    tSAMap.Bounds.cElements = CLng(pEndScan - pStartScan) + 1
    tSAMap.pvData = pStartScan

    DupArray bData, VarPtr(tSAMap)
    
    lIndex = FindSignature(bData(), bTemplate(), bMask())
    
    DupArray bData, 0@
    
    If lIndex = -1 Then Exit Function
    
    CopyMemory SearchEbMode, ByVal pStartScan + lIndex + &HA, Len(SearchEbMode)

#Else
    
    SearchEbMode = GetProcAddress(hVbe, "EbMode")
    
#End If

End Function

' // Create 32 bit thunk
#If VBA7 Then
Private Function Create32BitThunk() As LongPtr
    Dim pCode           As LongPtr
    Dim pfnKillTimer    As LongPtr
    Dim pfnTimerProc    As LongPtr
    Dim pVtbl           As LongPtr
    Dim hUser32         As LongPtr
#Else
Private Function Create32BitThunk() As Long
    Dim pCode           As Long
    Dim pfnKillTimer    As Long
    Dim pfnTimerProc    As Long
    Dim pVtbl           As Long
    Dim hUser32         As Long
#End If
    Dim bIsInIDE        As Boolean
    Dim llThunk(8)      As Currency
    
#If Not CBool(VBA6 Or VBA7) Then
    Debug.Assert MakeTrue(bIsInIDE)
#Else
    bIsInIDE = True
#End If

    If m_hCodeHeap = 0 Then Exit Function
    
    ' // Get TimerProc address
    CopyMemory pVtbl, ByVal ObjPtr(Me), Len(pVtbl)
    CopyMemory pfnTimerProc, ByVal pVtbl + (TIMERPROC_INDEX + 7) * Len(pfnTimerProc), Len(pfnTimerProc)

    If bIsInIDE Then
    
        hUser32 = GetModuleHandle(StrPtr("user32"))
        If hUser32 = 0 Then
        
            hUser32 = LoadLibrary(StrPtr("user32"))
            If hUser32 = 0 Then
                Exit Function
            End If
            
        End If
        
        pfnKillTimer = GetProcAddress(hUser32, "KillTimer")
        If pfnKillTimer = 0 Then Exit Function
        
        pCode = HeapAlloc(m_hCodeHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, &H48)
        If pCode = 0 Then Exit Function
        
        pCode = pCode + 4   ' // Disable
        
        ' // E8 00 00 00 00 85 C0 74 06 3C 01 74 16 EB 32 FF
        ' // 0D 00 00 00 00 68 00 00 00 00 6A 00 E8 3A 10 00
        ' // 00 EB 1E 6A 00 54 FF 74 24 18 FF 74 24 18 FF 74
        ' // 24 18 FF 74 24 18 68 00 00 00 00 E8 00 00 00 00
        ' // 58 C2 10 00
        
        llThunk(0) = 841287033897458.0968@
        llThunk(1) = -5772536353434.9306@
        llThunk(2) = 11434920928.8717@
        llThunk(3) = 456836774114.0992@
        llThunk(4) = 843054938821800.2176@
        llThunk(5) = 843048357232162.2052@
        llThunk(6) = 2929994243867.242@
        llThunk(7) = 389231.4112@
        llThunk(8) = 109.8328@
        
        CopyMemory ByVal pCode, llThunk(0), &H44
        CopyMemory ByVal pCode + 1, m_pEbMode - (pCode + 5), Len(m_pEbMode)
        CopyMemory ByVal pCode + &H11, pCode - 4, Len(pCode)
        CopyMemory ByVal pCode + &H1D, pfnKillTimer - (pCode + &H1C + 5), Len(pfnKillTimer)
        CopyMemory ByVal pCode + &H37, ObjPtr(Me), Len(pCode)
        CopyMemory ByVal pCode + &H3C, pfnTimerProc - (pCode + &H3B + 5), Len(pfnTimerProc)
    
    Else
        
        pCode = HeapAlloc(m_hCodeHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, &H29)
        If pCode = 0 Then Exit Function
        
        pCode = pCode + 4   ' // Disable
        
        llThunk(0) = -6526210967837.071@
        llThunk(1) = -6526210968229.3644@
        llThunk(2) = 174641.266@
        llThunk(3) = 120762440711195.876@
        
        CopyMemory ByVal pCode, llThunk(0), &H25
        CopyMemory ByVal pCode + &H14, ObjPtr(Me), Len(pCode)
        CopyMemory ByVal pCode + &H19, pfnTimerProc - (pCode + &H18 + 5), Len(pfnTimerProc)
        
    End If
    
    Create32BitThunk = pCode
    
End Function

#End If

' // Search for signature using mask
Private Function FindSignature( _
                 ByRef bData() As Byte, _
                 ByRef bSignature() As Byte, _
                 ByRef bMask() As Byte) As Long
    Dim lDataIndex  As Long
    Dim lSignIndex  As Long
    
    lDataIndex = 0: lSignIndex = 0
    
    Do While lDataIndex <= UBound(bData)

        If bData(lDataIndex) = bSignature(lSignIndex) Or bMask(lSignIndex) = 0 Then
        
            lSignIndex = lSignIndex + 1
            
            If lSignIndex > UBound(bSignature) Then
                FindSignature = lDataIndex - UBound(bSignature)
                Exit Function
            End If
            
        Else
        
            If lSignIndex Then
                lDataIndex = lDataIndex - lSignIndex + 1
                lSignIndex = 0
            End If
            
        End If
        
        lDataIndex = lDataIndex + 1
        
    Loop
    
    FindSignature = -1
    
End Function

' // Get heap for asm thunks
' // Class save it to global env variable
#If VBA7 Then
Private Function GetCodeHeap() As LongPtr
#Else
Private Function GetCodeHeap() As Long
#End If
    Dim sHeapHandleString   As String
    Dim lIndex              As Long
    
    If m_hCodeHeap Then
        GetCodeHeap = m_hCodeHeap
        Exit Function
    End If
    
    sHeapHandleString = Space$(Len(GetCodeHeap) * 2)
    
    If GetEnvironmentVariable(StrPtr(HEAP_ENV_VARIABLE), StrPtr(sHeapHandleString), LenB(sHeapHandleString)) Then
#If VBA7 Then
        m_hCodeHeap = CLngPtr("&H" & sHeapHandleString)
#Else
        m_hCodeHeap = CLng("&H" & sHeapHandleString)
#End If
        GetCodeHeap = m_hCodeHeap
        Exit Function
    End If
    
    m_hCodeHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, 0, 0)
    If m_hCodeHeap = 0 Then Exit Function
     
    sHeapHandleString = Hex$(m_hCodeHeap)
    
    For lIndex = Len(sHeapHandleString) + 1 To Len(GetCodeHeap) * 2
        sHeapHandleString = "0" & sHeapHandleString
    Next
    
    SetEnvironmentVariable StrPtr(HEAP_ENV_VARIABLE), StrPtr(sHeapHandleString)
    
    GetCodeHeap = m_hCodeHeap
    
End Function

' // Extract EbMode function from previous thunks
#If VBA7 Then
Private Function GetEbModeFromThunks() As LongPtr
#Else
Private Function GetEbModeFromThunks() As Long
#End If
    Dim tEntry  As PROCESS_HEAP_ENTRY
    
    If m_hCodeHeap = 0 Then Exit Function
    
    HeapLock m_hCodeHeap
    
    Do While HeapWalk(m_hCodeHeap, tEntry)

#If VBA7 And Win64 Then
        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY And _
            tEntry.cbData >= &H1E Then
            
            CopyMemory GetEbModeFromThunks, ByVal tEntry.lpData + &H1E, Len(GetEbModeFromThunks)
#Else
        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY And _
            tEntry.cbData >= &H1E Then
            
            CopyMemory GetEbModeFromThunks, ByVal tEntry.lpData + &H5, Len(GetEbModeFromThunks)
            
            GetEbModeFromThunks = GetEbModeFromThunks + (tEntry.lpData + &H4) + 5
            
#End If
            
            Exit Do
            
        End If
        
    Loop

    HeapUnlock m_hCodeHeap
    
End Function

' // Check if there is inactive thunks and free them
' // Returns the number of active thunks
Private Function CleanupThunks() As Long
    Dim tEntry      As PROCESS_HEAP_ENTRY
    Dim lDisable    As Long
    Dim lCount      As Long
#If VBA7 Then
    Dim pThunk      As LongPtr
#Else
    Dim pThunk      As Long
#End If

    If m_hCodeHeap = 0 Then Exit Function
    
    HeapLock m_hCodeHeap
    
    Do While HeapWalk(m_hCodeHeap, tEntry)
        
        If pThunk Then
            HeapFree m_hCodeHeap, HEAP_NO_SERIALIZE, pThunk
            pThunk = 0
        End If
        
        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY And _
            tEntry.cbData >= Len(lDisable) Then
            
            ' // Check if disabled
            CopyMemory lDisable, ByVal tEntry.lpData, Len(lDisable)
            
            If lDisable Then
                pThunk = tEntry.lpData
            Else
                lCount = lCount + 1
            End If
            
        End If
        
    Loop
    
    If pThunk Then
        HeapFree m_hCodeHeap, HEAP_NO_SERIALIZE, pThunk
    End If
        
    HeapUnlock m_hCodeHeap
    
    CleanupThunks = lCount
    
End Function

Private Sub Class_Terminate()
    
    If m_lIdEvent Then
        KillTimer 0, m_lIdEvent
        m_lIdEvent = 0
    End If
    
    If m_pAsmThunk Then
        HeapFree m_hCodeHeap, HEAP_NO_SERIALIZE, m_pAsmThunk - 4
        m_pAsmThunk = 0
    End If
    
    If CleanupThunks() = 0 Then
        
        HeapDestroy m_hCodeHeap
        m_hCodeHeap = 0
        SetEnvironmentVariable StrPtr(HEAP_ENV_VARIABLE), 0
        
    End If
    
End Sub

#If Not CBool(VBA6 Or VBA7) Then
Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function
#End If


