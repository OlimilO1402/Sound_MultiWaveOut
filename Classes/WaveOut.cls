VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public IsPlaying       As Boolean
Private mIsPaused      As Boolean
Private mWaveFmtEx     As TWaveFormatEx
Private mWaveHeader    As TWaveHeader
Private mWaveOutHandle As Long
Private mWaveOutDevice As Long
'Private mbolLoop       As Boolean

Private Sub Class_Terminate()
    'den Header unpreparen
    Call UnPrepareHeader
    'WaveOut wieder schließen
    Call CClose
End Sub

Private Function OOpen() As Long
    'gibt das WaveOut-Handle zurück
    Dim hr    As EMMResult
    'auch ein Device für den Handle öffnen
    Dim devid As Long
    For devid = 0 To waveOutGetNumDevs - 1
        hr = waveOutOpen(OOpen, _
                         EWaveOutDevice.WaveMapper, _
                         mWaveFmtEx, _
                         AddressOf CallBack_WaveOutProc, _
                         ObjPtr(Me), _
                         EWaveOutFlags.CallBackFunction)
        
        If OOpen = 0 Then
            Debug.Print "Handle not valid"
        Else
            'MsgBox "Handle OK: " & CStr(OOpen)
        End If
        If hr = NoError Then
            mWaveOutDevice = devid
            Exit For
        Else
            Debug.Print EMMResultToString(hr)
        End If
    Next
End Function

Public Sub PlayWaveSound(Wave As WaveSound)
    Dim hr As EMMResult
    'das WaveFormat übernehmen
    If Wave Is Nothing Then Exit Sub
    If IsPlaying Then Exit Sub
    'Evtl überprüfen ob das vorhergehende WaveFormat mit dem neuen übereinstimmt.
    'Falls nicht, so müßte man das WaveoutHandle schließen und ein Neues öffnen
    'evtl noch überprüfen ob das WaveFormat unterstützt wird
    'oder man macht diese Überprüfung schon vorher außerhalb der Klasse
    If mWaveFmtEx.AvgBytesPerSec = 0 Then
        LSet mWaveFmtEx = Wave.WaveFormat
    End If
    'den WaveOutKanal öffnen
    If mWaveOutHandle = 0 Then mWaveOutHandle = OOpen()
    If mWaveOutHandle = 0 Then Exit Sub 'Meldung wird schon in OOpen ausgegeben
    'evtl den alten Header löschen
    'den alten Header nicht löschen, wenn Neuer und Alter gleich sind
    If mWaveHeader.Flags And WhdrPrepared Then
        If Not IsEqualDataBuffer(mWaveHeader.lpData, mWaveHeader.BufferLength, _
                                 Wave.pData, Wave.ByteLength) Then
            Call UnPrepareHeader
            'den Header neu aufstellen
            hr = PrepareHeader(Wave)
        End If
    Else
        'den Header neu aufstellen
        hr = PrepareHeader(Wave)
    End If
    If hr = NoError Then
        Call Play
    End If
End Sub

Public Function Play() As EMMResult
    Play = waveOutWrite(mWaveOutHandle, mWaveHeader, LenB(mWaveHeader))
    mIsPaused = False
    If Play = NoError Then
        IsPlaying = True
    Else
        Debug.Print EMMResultToString(Play)
    End If
End Function

Public Sub Pause()
    Dim hr As EMMResult
    If Not mIsPaused Then hr = waveOutPause(mWaveOutHandle)
    mIsPaused = True
    If hr = NoError Then
        mIsPaused = True
    Else
        Debug.Print EMMResultToString(hr)
    End If
End Sub

Public Property Get IsPaused() As Boolean
    IsPaused = mIsPaused
End Property
Public Sub ReStart()
    Dim hr As EMMResult
    hr = waveOutRestart(mWaveOutHandle)
    If hr <> NoError Then Debug.Print EMMResultToString(hr)
End Sub

Public Sub SStop()
    '
End Sub

Public Function CClose() As EMMResult
    'Schließt den WaveOutKanal
    If mWaveOutHandle <> 0 Then
        CClose = waveOutClose(mWaveOutHandle)
        If CClose <> NoError Then Debug.Print EMMResultToString(CClose)
        mWaveOutHandle = 0
    End If
End Function

Public Function UnPrepareHeader() As EMMResult
    If mWaveOutHandle <> 0 Then
        UnPrepareHeader = waveOutUnprepareHeader(mWaveOutHandle, mWaveHeader, LenB(mWaveHeader))
        If UnPrepareHeader <> NoError Then Debug.Print EMMResultToString(UnPrepareHeader)
    End If
End Function

Private Function PrepareHeader(Wave As WaveSound) As EMMResult
    With Wave
        mWaveHeader = New_TWaveHeader(.pData, .ByteLength)
    End With
    PrepareHeader = waveOutPrepareHeader(mWaveOutHandle, mWaveHeader, LenB(mWaveHeader))
    If PrepareHeader <> NoError Then Debug.Print EMMResultToString(PrepareHeader)
End Function
