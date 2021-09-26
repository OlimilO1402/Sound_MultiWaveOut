Attribute VB_Name = "MMWaveOut"
Option Explicit
'Public Const MMSYSERR_NOERROR      As Long = 0  '  no error
'Public Const MMSYSERR_BASE         As Long = 0
'Public Const MMSYSERR_ERROR        As Long = (MMSYSERR_BASE + 1)  '  unspecified error
'Public Const MMSYSERR_BADDEVICEID  As Long = (MMSYSERR_BASE + 2)  '  device ID out of range
'Public Const MMSYSERR_NOTENABLED   As Long = (MMSYSERR_BASE + 3)  '  driver failed enable
'Public Const MMSYSERR_ALLOCATED    As Long = (MMSYSERR_BASE + 4)  '  device already allocated
'Public Const MMSYSERR_INVALHANDLE  As Long = (MMSYSERR_BASE + 5)  '  device handle is invalid
'Public Const MMSYSERR_NODRIVER     As Long = (MMSYSERR_BASE + 6)  '  no device driver present
'Public Const MMSYSERR_NOMEM        As Long = (MMSYSERR_BASE + 7)  '  memory allocation error
'Public Const MMSYSERR_NOTSUPPORTED As Long = (MMSYSERR_BASE + 8)  '  function isn't supported
'Public Const MMSYSERR_BADERRNUM    As Long = (MMSYSERR_BASE + 9)  '  error value out of range
'Public Const MMSYSERR_INVALFLAG    As Long = (MMSYSERR_BASE + 10) '  invalid flag passed
'Public Const MMSYSERR_INVALPARAM   As Long = (MMSYSERR_BASE + 11) '  invalid parameter passed
'Public Const MMSYSERR_HANDLEBUSY   As Long = (MMSYSERR_BASE + 12) '  handle being used
'Public Const MMSYSERR_INVALIDALIAS As Long = (MMSYSERR_BASE + 13) '  "Specified alias not found in WIN.INI
'Public Const MMSYSERR_LASTERROR    As Long = (MMSYSERR_BASE + 13) '  last error in range
'Public Const MMSYSERR_BADDB        As Long = (MMSYSERR_BASE + 14)
'Public Const MMSYSERR_KEYNOTFOUND  As Long = (MMSYSERR_BASE + 15)
'Public Const MMSYSERR_READERROR    As Long = (MMSYSERR_BASE + 16)
'Public Const MMSYSERR_WRITEERROR   As Long = (MMSYSERR_BASE + 17)
'Public Const MMSYSERR_DELETEERROR  As Long = (MMSYSERR_BASE + 18)
'Public Const MMSYSERR_VALNOTFOUND  As Long = (MMSYSERR_BASE + 19)
'Public Const MMSYSERR_NODRIVERCB   As Long = (MMSYSERR_BASE + 20)
'Public Const MMSYSERR_MOREDATA     As Long = (MMSYSERR_BASE + 21)
'Public Const WAVERR_BASE           As Long = 32
'Public Const WAVERR_BADFORMAT      As Long = (WAVERR_BASE + 0)
'Public Const WAVERR_STILLPLAYING   As Long = (WAVERR_BASE + 1)
'Public Const WAVERR_SYNC           As Long = (WAVERR_BASE + 3)
'Public Const WAVERR_LASTERROR      As Long = (WAVERR_BASE + 3)

Public Enum EMMResult
    NoError = 0
    AnyError = 1
    BadDeviceID = 2
    NotEnabled = 3
    Allocated = 4
    InvalidHandle = 5
    NoDriver = 6
    NoMem = 7
    NotSupported = 8
    BadErrNum = 9
    InvalidFlag = 10
    InvalidParam = 11
    HandleBusy = 12
    InvalidAlias = 13
    LastErr = 13
    BaddDB = 14
    KeyNotFound = 15
    ReadError = 16
    WriteError = 17
    DeleteError = 18
    ValNotFound = 19
    NoDriverCB = 20
    MoreData = 21
    'gap?
    BadFormat = 32
    StillPlaying = 33
    SyncError = 35
End Enum

'Public Const WAVE_MAPPER    As Long = -1&
Public Enum EWaveOutDevice
    WaveMapper = -1
    'What Else? No Idea
End Enum

'Konstanten für fdwOpen:
'Public Const CALLBACK_NULL      As Long = &H0     'No callback mechanism. This is the default setting.
'Public Const CALLBACK_WINDOW    As Long = &H10000 'The dwCallback parameter is a window handle.
'Public Const CALLBACK_THREAD    As Long = &H20000 'The dwCallback parameter is a thread identifier.
'Public Const CALLBACK_FUNCTION  As Long = &H30000 'The dwCallback parameter is a callback procedure address.
'Public Const CALLBACK_EVENT     As Long = &H50000 'The dwCallback parameter is an event handle.
'
'Public Const WAVE_FORMAT_QUERY  As Long = &H1     'If this flag is specified, waveOutOpen queries the device to
                                                   'determine if it supports the given format, but the device is not
                                                   'actually opened.
'Public Const WAVE_FORMAT_DIRECT As Long = &H2     'If this flag is specified, the ACM driver does not perform
                                                   'conversions on the audio data.
'Public Const WAVE_MAPPED        As Long = &H4     'If this flag is specified, the uDeviceID parameter specifies a
                                                   'waveform-audio device to be mapped to by the wave mapper.
'Public Const WAVE_ALLOWSYNC     As Long = &H10000 'If this flag is specified, a synchronous waveform-audio device can
                                                   'be opened. If this flag is not specified while opening a synchronous
                                                   'driver, the device will fail to open.

Public Enum EWaveOutFlags
    CallBackNull = 0
    CallBackWindow = &H10000
    CallBackThread = &H20000
    CallBackFunction = &H30000
    CallBackEvent = &H50000
    WaveFmtQuery = &H1
    WaveFmtDirect = &H2
    WaveFmtMapped = &H4
    WaveFmtAllowSync = &H10000
End Enum

'Public Const WHDR_DONE      As Long = &H1
'Public Const WHDR_PREPARED  As Long = &H2
'Public Const WHDR_BEGINLOOP As Long = &H4
'Public Const WHDR_ENDLOOP   As Long = &H8
'Public Const WHDR_INQUEUE   As Long = &H10
'Public Const WHDR_VALID     As Long = &H1F
'
Public Enum EWaveHeaderFlags
    WhdrNone = &H0&
    WhdrDone = &H1&
    WhdrPrepared = &H2&
    WhdrBeginLoop = &H4&
    WhdrEndLoop = &H8&
    WhdrInQueue = &H10&
    WhdrValid = &H1F&
End Enum

'typedef struct wavehdr_tag {
'    LPSTR      lpData;
'    DWORD      dwBufferLength;
'    DWORD      dwBytesRecorded;
'    DWORD_PTR  dwUser;
'    DWORD      dwFlags;
'    DWORD      dwLoops;
'    struct wavehdr_tag * lpNext;
'    DWORD_PTR reserved;
'} WAVEHDR, *LPWAVEHDR;

Public Type TWaveHeader 'WAVEHDR
    lpData        As Long
    BufferLength  As Long
    BytesRecorded As Long
    dwUser        As Long
    Flags         As EWaveHeaderFlags 'As Long
    Loops         As Long
    lpNext        As Long
    Reserved      As Long
End Type

' lpData          | Pointer to the waveform buffer.
' dwBufferLength  | Length, in bytes, of the buffer.
' dwBytesRecorded | When the header is used in input, this member specifies
'                 | how much data is in the buffer.
' dwUser          | User data.
' dwFlags         | Flags supplying information about the buffer.
'                 | The following values are defined:
' WHDR_BEGINLOOP  | This buffer is the first buffer in a loop.
'                 | This flag is used only with output buffers.
' WHDR_DONE       | Set by the device driver to indicate that it is finished with the
'                 | buffer and is returning it to the application.
'                 | WHDR_ENDLOOP
'                 | This buffer is the last buffer in a loop. This flag is used only
'                 | with output buffers.
' WHDR_INQUEUE    | Set by Windows to indicate that the buffer is queued for playback.
' WHDR_PREPARED   | Set by Windows to indicate that the buffer has been prepared with
'                 | the wavInPrepareHeader or waveOutPrepareHeader function.
' dwLoops         | Number of times to play the loop. This member is used only
'                 | with output buffers.
' lpNext          | Reserved.
' reserved        | Reserved.
'
' Remarks
' =======
' Use the WHDR_BEGINLOOP and WHDR_ENDLOOP flags in the dwFlags member to specify
' the beginning and ending data blocks for looping. To loop on a single block,
' specify both flags for the same block. Use the dwLoops member in the WAVEHDR
' structure for the first block in the loop to specify the number of times to
' play the loop.
' The lpData, dwBufferLength, and dwFlags members must be set before calling the
' waveInPrepareHeader or waveOutPrepareHeader function. (For either function, the
' dwFlags member must be set to zero.)

Public Const MM_WOM_OPEN  As Long = &H3BB
Public Const MM_WOM_DONE  As Long = &H3BD
Public Const MM_WOM_CLOSE As Long = &H3BC

Public Const WOM_OPEN     As Long = MM_WOM_OPEN
Public Const WOM_DONE     As Long = MM_WOM_DONE
Public Const WOM_CLOSE    As Long = MM_WOM_CLOSE

'MMRESULT waveOutOpen(
'  LPHWAVEOUT     phwo,
'  UINT_PTR       uDeviceID,
'  LPWAVEFORMATEX pwfx,
'  DWORD_PTR      dwCallback,
'  DWORD_PTR      dwCallbackInstance,
'  DWORD fdwOpen
');
'die Funktion waveOutOpen übergibt byref das WaveOutHandle!
'die Übergabe des WaveFormats ByRef As Any hat den Vorteil
'daß man sowohl die Variable ByRef als auch einen Zeiger ByVal
'darauf übergeben kann.
Public Declare Function waveOutOpen Lib "winmm" ( _
                 ByRef lphWaveOut As Long, _
                 ByVal uDeviceID As Long, _
                 ByRef pwfmtx As Any, _
                 ByVal dwCallback As Long, _
                 ByVal dwCallbackInstance As Long, _
                 ByVal dwFlags As EWaveOutFlags _
                 ) As EMMResult 'Long

Public Declare Function waveOutClose Lib "winmm" ( _
                 ByVal hWaveOut As Long _
                 ) As EMMResult 'Long

Public Declare Function waveOutWrite Lib "winmm" ( _
                 ByVal hWaveOut As Long, _
                 ByRef lpWaveOutHdr As Any, _
                 ByVal uSize As Long _
                 ) As EMMResult 'Long
                 
Public Declare Function waveOutPrepareHeader Lib "winmm" ( _
                 ByVal hWaveOut As Long, _
                 ByRef lpWaveOutHdr As Any, _
                 ByVal uSize As Long _
                 ) As EMMResult 'Long

Public Declare Function waveOutUnprepareHeader Lib "winmm" ( _
                 ByVal hWaveOut As Long, _
                 ByRef lpWaveOutHdr As Any, _
                 ByVal uSize As Long _
                 ) As EMMResult 'Long

Public Declare Function waveOutGetNumDevs Lib "winmm" ( _
                 ) As Long

Public Declare Function waveOutPause Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long _
                 ) As Long

Public Declare Function waveOutGetPosition Lib "winmm" ( _
                 ByVal hWaveOut As Long, _
                 lpInfo As Any, _
                 ByVal uSize As Long _
                 ) As Long

Public Declare Function waveOutRestart Lib "winmm" ( _
                 ByVal hWaveOut As Long _
                 ) As Long

Public Function Min(val1, val2)
    If val1 < val2 Then Min = val1 Else Min = val2
End Function

Public Function EMMResultToString(hr As EMMResult) As String
    Dim s As String
    Select Case hr
    Case NoError:       s = "NoError"
    Case AnyError:      s = "AnyError"
    Case BadDeviceID:   s = "BadDeviceID"
    Case NotEnabled:    s = "NotEnabled"
    Case Allocated:     s = "Allocated"
    Case InvalidHandle: s = "InvalidHandle"
    Case NoDriver:      s = "NoDriver"
    Case NoMem:         s = "NoMem"
    Case NotSupported:  s = "NotSupported"
    Case BadErrNum:     s = "BadErrNum"
    Case InvalidFlag:   s = "InvalidFlag"
    Case InvalidParam:  s = "InvalidParam"
    Case HandleBusy:    s = "HandleBusy"
    Case InvalidAlias:  s = "InvalidAlias"
    'Case LastErr:
    Case BaddDB:        s = "BaddDB"
    Case KeyNotFound:   s = "KeyNotFound"
    Case ReadError:     s = "ReadError"
    Case WriteError:    s = "WriteError"
    Case DeleteError:   s = "DeleteError"
    Case ValNotFound:   s = "ValNotFound"
    Case NoDriverCB:    s = "NoDriverCB"
    Case MoreData:      s = "MoreData"
    'gap?
    Case BadFormat:     s = "BadWaveFormat"
    Case StillPlaying:  s = "StillPlaying"
    Case SyncError:     s = "SyncError"
    Case Else: s = "undefined error: " & CStr(hr) & " &H" & Hex$(hr)
    End Select
    EMMResultToString = s
End Function

Public Function EWaveHeaderFlagsToString(this As EWaveHeaderFlags) As String
    Dim s As String
    If this And WhdrNone Then AddAnd(s) = "WhdrNone"
    If this And WhdrDone Then AddAnd(s) = "WhdrDone"
    If this And WhdrPrepared Then AddAnd(s) = "WhdrPrepared"
    If this And WhdrBeginLoop Then AddAnd(s) = "WhdrBeginLoop"
    If this And WhdrEndLoop Then AddAnd(s) = "WhdrEndLoop"
    If this And WhdrInQueue Then AddAnd(s) = "WhdrInQueue"
    If this And WhdrValid Then AddAnd(s) = "WhdrValid"
    EWaveHeaderFlagsToString = s
End Function
Private Property Let AddAnd(this As String, s As String)
    If Len(this) Then this = this & " And " & s Else this = this & s
End Property
Public Function FncPtr(pFnc As Long) As Long
    FncPtr = pFnc
End Function

Public Function New_TWaveHeader(ByVal pData As Long, _
                                ByVal Length As Long, _
                                Optional ByVal dwUser As Long, _
                                Optional ByVal Flags As EWaveHeaderFlags, _
                                Optional ByVal Loops As Long = 1, _
                                Optional ByVal pNext As Long _
                                ) As TWaveHeader
    With New_TWaveHeader
        .lpData = pData
        .BufferLength = Length
        '.BytesRecorded = 0
        .dwUser = dwUser
        .Flags = Flags
        .Loops = Loops
        .lpNext = pNext
        '.reserved
    End With
End Function
Public Function IsEqualDataBuffer(lpData1 As Long, BufLen1 As Long, lpData2 As Long, BufLen2 As Long) As Boolean
    If lpData1 <> lpData2 Then Exit Function
    If BufLen1 <> BufLen2 Then Exit Function
    IsEqualDataBuffer = True
End Function
'Public Function IsEqualWaveHeader(whdr1 As TWaveHeader, whdr2 As TWaveHeader) As Boolean
'    If whdr1.lpData <> whdr2.lpData Then Exit Function
'    If whdr1.BufferLength <> whdr2.BufferLength Then Exit Function
'    If whdr1.BytesRecorded <> whdr2.BytesRecorded Then Exit Function
'    If whdr1.dwUser <> whdr2.dwUser Then Exit Function
'    If whdr1.Flags <> whdr2.Flags Then Exit Function
'    If whdr1.Loops <> whdr2.Loops Then Exit Function
'    If whdr1.lpNext <> whdr2.lpNext Then Exit Function
'    IsEqualWaveHeader = True
'End Function
Public Function WaveOutProcMsgToString(ByVal c As Long) As String
    Select Case c
    Case WOM_OPEN:  WaveOutProcMsgToString = "WOM_OPEN"
    Case WOM_DONE:  WaveOutProcMsgToString = "WOM_DONE"
    Case WOM_CLOSE: WaveOutProcMsgToString = "WOM_CLOSE"
    Case Else:      WaveOutProcMsgToString = "WOM_ELSE" 'What else? nothing else!
    End Select
End Function

Public Sub CallBack_WaveOutProc(ByVal hwo As Long, _
                                ByVal uMsg As Long, _
                                ByVal aWaveOut As WaveOut, _
                                ByVal pWaveHeader As Long, _
                                ByVal dwParam2 As Long)
    'hier besser die Debug-Geschichten sein lassen
    'DebugPrint "CallBack_WaveOutProc " & WaveOutProcMsgToString(uMsg)
    Select Case uMsg
    Case WOM_OPEN:  'nichts machen, kann man auch entfernen, ist nur der Vollständigkeit halber drin
    Case WOM_DONE
        'hier in der CallBack-Prozedur (als CALLBACK_FUNCTION) dürfen keine WaveOut-Funktionen
        'aufgerufen werden. Die möglichen übergebbaren Daten (WaveoutHandle, WaveHeader etc.)
        'können damit leider garnicht verwendet werden. Das einzige was man machen kann ist dem
        'WaveOut-Objekt mitzuteilen, daß der Buffer fertig gespielt hat.
        'Dies ist allerdings nicht weiter schlimm, da man um zeitgesteuert mehrere Sounds auszugeben
        'ohnehin einen Timer verwendet, und die Funktion UnprepareHeader auch zu einem späteren
        'Zeitpunkt aufgerufen werden kann.
        If Not aWaveOut Is Nothing Then
            aWaveOut.IsPlaying = False
        End If
    Case WOM_CLOSE: 'nichts machen, kann man auch entfernen, ist nur der Vollständigkeit halber drin
    End Select
End Sub

Public Sub DebugPrint(DbgMsg As String, Optional Clear As Boolean)
    
    'DbgMsg : die Meldung die ausgegeben werden soll
    'falls Clear = True, dann
    '    sollen alle bestehenden Label.Caption = vbNullString
    '    DebugCount = 0
    'Label werden nur nachgeladen,
    'Label werden nicht gelöscht
    'DebugCount: die Anzahl an bisherigen Meldungen
    'Achtung hier Baustelle!!!
    '
    'c: die Anzahl an Label die bereits geladen sind
    Dim c As Long
    Dim i As Long
    
    Dim Label 'As Label
    Set Label = Form1.Label1
    
    c = Label.Count
    If c = 1 Then Label(0).AutoSize = True
    
    Static DebugCount As Long
    DebugCount = IIf(Clear, 1, DebugCount + 1)
    
    If DebugCount < c Then
        For i = DebugCount To c - 1
            Label(i).Caption = vbNullString
        Next
    Else
        'das Label für den nachfoldenen schon laden
        'dieses Label wird aber jetzt noch nicht benützt
        Load Label(c)
        'das neue Label positionieren
        Label(c).Top = Label(c - 1).Top + 240
    End If
    Label(DebugCount - 1).Visible = True
    Label(DebugCount - 1).Caption = DbgMsg
    '
    'hat die Meldung keine Länge, dann den auch wieder freigeben
    '
    If Len(DbgMsg) = 0 Then DebugCount = DebugCount - 1
    Debug.Print DbgMsg
    
End Sub
