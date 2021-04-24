VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mWaveFormat As TWaveFormat
Private mWaveData() As Byte

Public Sub NewC(WaveFileName As String)
    Call Load(WaveFileName)
End Sub

Public Function Load(WaveFileName As String) As Boolean
    Dim FNr        As Integer
    Dim WavefmtPos As Long
    Dim DataLenPos As Long
    Dim DataLength As Long
    Dim RiffLength As Long
TryE: On Error GoTo CatchE
    FNr = FreeFile
    Open WaveFileName For Binary Access Read As FNr
    Get FNr, 5, RiffLength
    
    WavefmtPos = InStrFile(FNr, C_WAVEfmt_)
    
    If WavefmtPos = 0 Then Exit Function
    DataLenPos = InStrFile(FNr, C_data, WavefmtPos + LenB(mWaveFormat))
    
    If DataLenPos = 0 Then Exit Function
    Get FNr, WavefmtPos + 12, mWaveFormat
    Get FNr, DataLenPos + 4, DataLength
    ReDim mWaveData(0 To (DataLength) - 1)
    Get FNr, , mWaveData
    Load = True
    'Debug.Print TWaveFormatToString(mWaveFormat)
FinallyE:
    Close FNr
    Exit Function
CatchE:
    MsgBox Err.Number & " " & Err.Description
    GoTo FinallyE
End Function

Public Property Get pData() As Long
    pData = VarPtr(mWaveData(0)) '- 1 + mDataPos
End Property

'whdr.lpData should contain a pointer to the wave data itself, and
'whdr.dwBufferLength the lengh of that wave data alone
Public Property Get ByteLength() As Long
    ByteLength = UBound(mWaveData) + 1
End Property

'ReadOnly
Friend Property Get WaveFormat() As TWaveFormat
    WaveFormat = mWaveFormat
End Property

Public Function LengthMS() As Long
'liefert die zeitliche Länge in Millisekunden
    LengthMS = ByteLength / mWaveFormat.AvgBytesPerSec * 1000
End Function

