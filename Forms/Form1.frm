VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox ChkChord 
      Caption         =   "chord.wav"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Value           =   1  'Aktiviert
      Width           =   1215
   End
   Begin VB.CheckBox ChkChimes 
      Caption         =   "chimes.wav"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1215
   End
   Begin VB.CheckBox ChkTada 
      Caption         =   "tada.wav"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton BtnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton BtnPause 
      Caption         =   "Pause"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton BtnStandingOvations 
      Caption         =   "StandingOvations"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnPlay32Channels 
      Caption         =   "Play32Channels"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnPlay123_2 
      Caption         =   "Play123_2"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnPlay123_1 
      Caption         =   "Play123_1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'im Normalfall würde man einen Timer verwenden
'wir verwenden hier mal der Einfachheit halbe Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwms As Long)
Private mWavOutChan1 As WaveOut
Private mWavOutChan2 As WaveOut
Private mWavOutChan3 As WaveOut

Private mMultiWave   As MultiWaveOut

Private mWavChimes   As WaveSound
Private mWavChord    As WaveSound
Private mWavTada     As WaveSound
Private mWavApplause As WaveSound

Private Const MediaPath As String = "C:\Windows\Media\"

Private Sub BtnPause_Click()
    Call mMultiWave.Pause
End Sub

Private Sub Form_Load()
    
    Randomize
    Set mWavOutChan1 = New WaveOut
    Set mWavOutChan2 = New WaveOut
    Set mWavOutChan3 = New WaveOut
    Set mWavChimes = New_WaveSound(MediaPath & "chimes.wav")
    Set mWavChord = New_WaveSound(MediaPath & "chord.wav")
    Set mWavTada = New_WaveSound(MediaPath & "tada.wav")
    Set mWavApplause = New_WaveSound(App.Path & "\Resources\Applause.wav")

End Sub

Private Sub BtnPlay123_1_Click()
    Set mMultiWave = New MultiWaveOut
    DebugPrint "BtnPlay123_1_Click", True
    'jetzt alle Kanäle die gleichzeitg gespielt werden sollen aufrufen
    If ChkChimes.Value Then Call mWavOutChan1.PlayWaveSound(mWavChimes)
    If ChkChord.Value Then Call mWavOutChan2.PlayWaveSound(mWavChord)
    If ChkTada.Value Then Call mWavOutChan3.PlayWaveSound(mWavTada)
    
    'mWavSnd3 wird nicht gespielt, da Kanal 2 noch am spielen ist!
    
End Sub

Private Sub BtnPlay123_2_Click()
    DebugPrint "BtnPlay123_2_Click", True
    'Man kann die Sounds aber auch so aufrufen. Jetzt wird
    'immer der Kanal gespielt, der gerade frei ist, wobei
    'der Aufruf in einer Zeile direkt sichtbar macht, daß
    'die drei Waves gleichzeitig wiedergegeben werden.
    Dim s As Long
    s = 1000
    Call mMultiWave.PlayMultiWave(mWavChimes, mWavChord, mWavTada)
    Call Sleep(s): DoEvents
    Call mMultiWave.PlayMultiWave(mWavChimes, mWavTada, mWavChord)
    Call Sleep(s): DoEvents
    Call mMultiWave.PlayMultiWave(mWavChord, mWavChimes, mWavTada)
    Call Sleep(s): DoEvents
    Call mMultiWave.PlayMultiWave(mWavChord, mWavTada, mWavChimes)
    Call Sleep(s): DoEvents
    Call mMultiWave.PlayMultiWave(mWavTada, mWavChimes, mWavChord)
    Call Sleep(s): DoEvents
    Call mMultiWave.PlayMultiWave(mWavTada, mWavChord, mWavChimes)

End Sub

Private Sub BtnPlay32Channels_Click()
    DebugPrint "BtnPlay32Channels_Click", True
    'hier wird jeder der 32 WaveOutKanäle nacheinander bemüht
    Dim i As Long
    Dim n As Long
    Dim t As Long
    Dim mess As String
    t = 125 '250
    n = mMultiWave.Count
    mess = "Jetzt wird jeder der " & CStr(n) & " Kanäle gespielt." & vbCrLf & _
           "Achtung das Folgende dauert ca " & CStr(n) * t / 1000 & " Sek."
    If MsgBox(mess, vbOKCancel) = vbCancel Then Exit Sub
    For i = 0 To n - 1
        Call mMultiWave.PlayChannel(i, mWavChimes)
        Call mMultiWave.PlayChannel(i, mWavChord)
        Call mMultiWave.PlayChannel(i, mWavTada)
        Call Sleep(t)
        DoEvents
    Next
End Sub

Private Sub BtnStandingOvations_Click()
    DebugPrint "BtnStandingOvations_Click", True
    Dim i As Long
    Dim l As Long
    Dim t As Long
    'Dim Applause As WaveSound: Set Applause = New_WaveSound(App.Path & "\Applause.wav")
    'Call mWavOutChan1.PlayWaveSound(Applause)
    'Call Sleep(1000)
    Set mMultiWave = New MultiWaveOut
    l = mWavApplause.LengthMS
    For i = 0 To 9
        Call mMultiWave.PlayMultiWave(mWavApplause)
        t = CLng(Rnd * 2000) 'CDbl(l))
        DebugPrint "Step: " & CStr(i) & ";  Waiting: " & CStr(t) & "ms"
        Call Sleep(t)
    Next
    'noch warten bis der letzte Applause beendet ist.
    t = l - t
    DebugPrint "Waiting: " & CStr(t) & "ms"
    Call Sleep(t)
End Sub

Private Sub BtnClose_Click()
    Call mWavOutChan1.CClose
    Call mWavOutChan2.CClose
    Call mWavOutChan3.CClose
    Call mMultiWave.CClose
End Sub

