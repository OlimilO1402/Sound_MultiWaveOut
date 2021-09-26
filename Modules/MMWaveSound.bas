Attribute VB_Name = "MMWaveSound"
Option Explicit

Public Const C_RIFF     As String = "RIFF"
Public Const C_WAVEfmt_ As String = "WAVEfmt "
Public Const C_data     As String = "data"

Public Const WAVE_FORMAT_PCM As Integer = 1

Public Type TWaveFormat
    FormatTag      As Integer
    Channels       As Integer
    SamplesPerSec  As Long
    AvgBytesPerSec As Long
    BlockAlign     As Integer
    BitsPerSample  As Integer
End Type

Public Type TWaveFormatEx
    FormatTag      As Integer ' 2
    Channels       As Integer ' 2
    SamplesPerSec  As Long    ' 4
    AvgBytesPerSec As Long    ' 4
    BlockAlign     As Integer ' 2
    BitsPerSample  As Integer ' 2
    ExtraDataSize  As Integer ' 2
End Type                 'Sum: 18 warum eigentlich nicht 4 aligned?

' FormatTag      | Format type. The following type is defined:
'                | WAVE_FORMAT_PCM Waveform-audio data is PCM.
' Channels       | Number of channels in the waveform-audio data. Mono data
'                | uses one channel and stereo data uses two channels.
' SamplesPerSec  | Sample rate, in samples per second.
' AvgBytesPerSec | Required average data transfer rate, in bytes per second.
'                | For example, 16-bit stereo at 44.1 kHz has an average data
'                | rate of 176,400 bytes per second (2 channels  — 2 bytes
'                | per sample per channel  — 44,100 samples per second).
' BlockAlign     | Block alignment, in bytes. The block alignment is the minimum
'                | atomic unit of data. For PCM data, the block alignment is the
'                | number of bytes used by a single sample, including data for
'                | both channels if the data is stereo. For example, the block
'                | alignment for 16-bit stereo PCM is 4 bytes
'                | (2 channels  — 2bytes per sample).
' BitsPerSample  | Bits per sample for the wFormatTag format type. If wFormatTag is
'                | WAVE_FORMAT_PCM, then wBitsPerSample should be equal to 8 or 16.
'                | For non-PCM formats, this member must be set according to the
'                | manufacturer's specification of the format tag. If wFormatTag is
'                | WAVE_FORMAT_EXTENSIBLE, this value can be any integer multiple of
'                | 8 and represents the container size, not necessarily the sample size;
'                | for example, a 20-bit sample size is in a 24-bit container. Some
'                | compression schemes cannot define a value for wBitsPerSample, so this
'                | member can be 0.
' ExtraDataSize  | Size, in bytes, of extra format information appended to the end of the
'                | WAVEFORMATEX structure. This information can be used by non-PCM formats
'                | to store extra attributes for the wFormatTag. If no extra information is
'                | required by the wFormatTag, this member must be set to 0.
'                | For WAVE_FORMAT_PCM formats (and only WAVE_FORMAT_PCM formats), this
'                | member is ignored. When this structure is included in a
'                | WAVEFORMATEXTENSIBLE structure, this value must be at least 22.


Public Function New_TWaveFormat(ByVal BitsPerSample As Long, _
                                ByVal SampleFrequency As Long, _
                                ByVal Channels As Integer) As TWaveFormat
    With New_TWaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = Channels
        .SamplesPerSec = SampleFrequency
        .BitsPerSample = BitsPerSample
        .BlockAlign = BitsPerSample / 8 * Channels
        .AvgBytesPerSec = SampleFrequency * .BlockAlign
    End With
End Function

Public Function TWaveFormatToString(this As TWaveFormat) As String
    Dim s As String
    With this
        s = s & "FormatTag: " & CStr(.FormatTag) & vbCrLf
        s = s & "Channels: " & CStr(.Channels) & vbCrLf
        s = s & "SamplesPerSec: " & CStr(.SamplesPerSec) & vbCrLf
        s = s & "AvgBytesPerSec: " & CStr(.AvgBytesPerSec) & vbCrLf
        s = s & "BlockAlign: " & CStr(.BlockAlign) & vbCrLf
        s = s & "BitsPerSample: " & CStr(.BitsPerSample) & vbCrLf
    End With
    TWaveFormatToString = s
End Function

Public Function WaveSound(WaveFileName As String) As WaveSound
    Set WaveSound = New WaveSound: WaveSound.New_ WaveFileName
End Function

Public Function InStrFile(ByVal FNr As Integer, _
                           StrSearch As String, _
                           Optional ByVal start As Long = 1) As Long

    ' liefert die Position eines Strings in einer Datei zurück
    Const C_LookupBuffLen As Long = 100

    Dim LookUpBuffer As String * C_LookupBuffLen
    Dim SearchLen    As Long
    Dim FileLength   As Long
    Dim i            As Long

    SearchLen = Len(StrSearch)
    FileLength = LOF(FNr)

    If (SearchLen / 2 > C_LookupBuffLen) Or (SearchLen > FileLength) Then Exit Function

    For i = start To ((FileLength / (C_LookupBuffLen / 2)) - start + 1) Step (C_LookupBuffLen / 2)
        Get FNr, i, LookUpBuffer
        InStrFile = InStr(1, LookUpBuffer, StrSearch, vbTextCompare)
        If InStrFile > 0 Then
            InStrFile = InStrFile + i - 1
            Exit For
        End If
    Next

End Function



