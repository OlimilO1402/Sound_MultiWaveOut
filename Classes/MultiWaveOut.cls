VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiWaveOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse verwaltet 32 WaveOut-Objekte in einem Array
Private Const C_CountMax                 As Long = 32
Private mWOChannels(0 To C_CountMax - 1) As New WaveOut

Public Sub PlayMultiWave(ParamArray waves() As Variant) 'As WaveSound) !
    'spielt immer den gerade freien Kanal
    Dim i  As Long
    Dim j  As Long
    Dim n  As Long
    Dim ws As WaveSound
    n = C_CountMax - 1
    If IsArray(waves) Then
        For i = LBound(waves) To Min(UBound(waves), n)
            Set ws = waves(i)
            For j = 0 To n
                With mWOChannels(j)
                    If Not .IsPlaying Then
                        Call .PlayWaveSound(ws)
                        DebugPrint "Now playing ChannelNr: " & CStr(j)
                        Exit For
                    End If
                End With
            Next
        Next
    End If
End Sub

Public Sub PlayChannel(Index As Long, aWave As WaveSound)
    If Index < C_CountMax Then
        Call mWOChannels(Index).PlayWaveSound(aWave)
        DebugPrint "Now playing ChannelNr: " & CStr(Index)
    End If
End Sub
Public Sub Pause()
    Dim i As Long
    For i = 0 To C_CountMax - 1
        With mWOChannels(i)
            If .IsPlaying Then .Pause
        End With
    Next
End Sub
Public Sub SStop()
    '
End Sub

Public Property Get Count() As Long
    Count = C_CountMax
End Property

Public Sub CCloseChannel(Index As Long)
    Call mWOChannels(Index).CClose
End Sub
Public Sub CClose()
    Dim i As Long
    For i = 0 To C_CountMax - 1
        Call mWOChannels(i).CClose
    Next
End Sub
