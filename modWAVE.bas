Attribute VB_Name = "modWAVE"

'Create the DirectSound8 Object
Public DX              As New DirectX8
Public ds              As DirectSound8



Private Const Pi = 3.141592654
Private Const Pi2 = 6.283185308

Public SampleRate      As Long    '11025

Public Sample          As Integer

Public SampleL         As Long    'Integer
Public SampleR         As Long    'Integer


Public Const Amplitude As Single = 30000    '127


Public Channels        As Integer



Public BitsPerSample   As Integer


Public WFfrom          As WAVEFORMATEX
Public WFto            As WAVEFORMATEX

Public BU()            As Byte


Public InpSound()      As Integer



Public bSize           As Long
Public CAP             As DSBCAPS



Public Sub makeFile(FileName As String, Optional SRate = 8000, Optional nChan = 1, Optional BitXsample = 16)
    SampleRate = SRate
    Channels = nChan
    BitsPerSample = BitXsample


    'FileName = App.Path & "\" & FileName
    On Error Resume Next
    Kill FileName            'REM this line if file does not exist

    Open FileName For Binary Access Write As #1
    Put #1, 1, "RIFF"        '"RIFF" header
    Put #1, 5, CInt(0)       'Filesize - 8, will write later
    Put #1, 9, "WAVEfmt "    '"WAVEfmt " header - not space after fmt
    Put #1, 17, CLng(16)     'Lenth of format data
    Put #1, 21, CInt(1)      'Wave type PCM
    Put #1, 23, CInt(Channels)    '1 channel
    Put #1, 25, CLng(SampleRate)    '44.1 kHz SampleRate
    Put #1, 29, CLng((SampleRate * BitsPerSample * Channels) / 8)
    Put #1, 33, CInt((BitsPerSample * Channels) / 8)
    Put #1, 35, CInt(BitsPerSample)
    Put #1, 37, "data"       '"data" Chunkheader
    Put #1, 41, CInt(0)      'Filesize - 44, will write later

End Sub
'Get the file length, write it into the header and close the file.
Public Sub closeFile()

    fileSize = LOF(1)
    Put #1, 5, CLng(fileSize - 8)
    Put #1, 41, CLng(fileSize - 44)
    Close #1


End Sub


'Define the DirectSound8 buffer, create it and set the play mode
Public Sub LoadAndPlayEngine(FileName As String, wBuff)

    Dim bufferDesc     As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS

    ' bufferDesc.fxFormat.nChannels = Channels
    'Stop

    '    Set ENGINEsound(wBuff) = ds.CreateSoundBufferFromFile(FileName, bufferDesc)
    Set B(wBuff).ENGINEsound = ds.CreateSoundBufferFromFile(FileName, bufferDesc)


    'ENGINEsound(wBuff).LoadAndPlayEngine DSBPLAY_LOOPING
    B(wBuff).ENGINEsound.Play DSBPLAY_LOOPING

    'ENGINEsound.LoadAndPlayEngine DSBPLAY_DEFAULT

    'ENGINEsound.SetFrequency 800

End Sub
Public Sub LoadDRIFT(FileName As String, wBuff)
'Stop

    Dim bufferDesc     As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS

    ' bufferDesc.fxFormat.nChannels = Channels
    'Stop

    '    Set DRIFTsound(wBuff) = ds.CreateSoundBufferFromFile(FileName, bufferDesc)
    Set B(wBuff).DRIFTsound = ds.CreateSoundBufferFromFile(FileName, bufferDesc)


    'B(wBuff).DRIFTsound.Play DSBPLAY_LOOPING
    'B(wBuff).DRIFTsound.Play DSBPLAY_DEFAULT


End Sub

Public Sub INITSound(formhWnd As Long, Nsound)

    On Local Error Resume Next
    Set ds = DX.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel formhWnd, DSSCL_NORMAL    'DSSCL_PRIORITY


    ' ReDim ENGINEsound(1 To Nsound)

End Sub


'Dispose of the DirectSound Object and its buffer
Public Sub CleanupSounds()

'If Not (ENGINEsound Is Nothing) Then ENGINEsound.Stop
'Set ENGINEsound = Nothing
    Set ds = Nothing
    Set DX = Nothing

End Sub

Public Sub StopSound(wS)

' ENGINEsound(wS).Stop
' Set ENGINEsound(wS) = Nothing
    B(wS).ENGINEsound.Stop
    B(wS).ENGINEsound = Nothing
End Sub


'Public Function LoadWaveAndConvert(inFileName, SampleRate, Bits, Channels) As Integer()'

'Dim bufferDesc As DSBUFFERDESC

'bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS


'Set ENGINEsound = ds.CreateSoundBufferFromFile(inFileName, bufferDesc)

'ENGINEsound.GetFormat WFfrom

'ENGINEsound.GetCaps CAP

'bSize = CAP.lBufferBytes

'ReDim BU(0 To bSize)

'ENGINEsound.ReadBuffer 0, bSize, BU(0), DSBLOCK_DEFAULT


'With WFfrom

'''''    MsgBox "Input File" & vbCrLf & "Sample Rate " & .lSamplesPerSec & " BitsPerSample " & .nBitsPerSample & " Channels " & .nChannels

'End With


'With WFto

'    .lSamplesPerSec = SampleRate 'IIf(WFfrom.lSamplesPerSec > 12000, WFfrom.lSamplesPerSec / 2, WFfrom.lSamplesPerSec)
'    .nBitsPerSample = Bits
'    .nChannels = Channels

'    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
'    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
'    .nFormatTag = WAVE_FORMAT_PCM

'End With

'ENGINEsound.SaveToFile App.Path & "\dsBufferOUT.wav"


''''InpSound = ConvertWave(BU, WFfrom, WFto)

'**************'LoadWaveAndConvert = ConvertWave(BU, WFfrom, WFto)

'End Function


Public Sub SaveArrayAsWave(FileName As String, Ar() As Integer, SamplesPerSec, Bits, Channels)


    Dim TmpWF          As WAVEFORMATEX
    Dim dsBuffer2      As DirectSoundSecondaryBuffer8


    With TmpWF
        .lSamplesPerSec = SamplesPerSec
        .nBitsPerSample = Bits
        .nChannels = Channels
        .nFormatTag = WAVE_FORMAT_PCM
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign

    End With


    Dim bufferDesc     As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS

    bufferDesc.lBufferBytes = (UBound(Ar)) * Channels
    bufferDesc.fxFormat = TmpWF

    Set dsBuffer2 = ds.CreateSoundBuffer(bufferDesc)


    dsBuffer2.WriteBuffer 0, UBound(Ar) * 2, Ar(0), DSBLOCK_DEFAULT

    dsBuffer2.SaveToFile FileName

End Sub


Public Function HexToDec(S As String)
    Dim CH1            As Integer
    Dim CH2            As Integer

    s1 = Left$(S, 1)
    S2 = Right$(S, 1)

    CH1 = Asc(s1)
    CH2 = Asc(S2)

    If CH1 > 58 Then
        CH1 = CH1 - 55
    Else
        CH1 = CH1 - 48
    End If
    If CH2 > 58 Then
        CH2 = CH2 - 55
    Else
        CH2 = CH2 - 48
    End If

    HexToDec = CH2 + CH1 * 16


End Function

