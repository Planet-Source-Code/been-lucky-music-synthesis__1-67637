VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSynthMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'WORKAROUND TO PRESERVE CASE OF ENUMS BELOW
'TIP CREDIT :Ulli's Enumeration Spelling. txtCodeId=42088

#If False Then
Private PureSine, WhiteNoise, SinePlusHarmonics, _
        SinePlusNoise, Sawtooth, SquareWave, _
        Other1, Other2, Other3, Other4, Other5
#End If

Public Enum SingleSample
    PureSine
    WhiteNoise
    SinePlusNoise
    SinePlusHarmonics
    Sawtooth
    SquareWave
    Other1
    Other2
    Other3
    Other4
    Other5
    'Other = 99
    End Enum
'--
Const Pi = 3.14159265358979 ' = 4 * Atn(1) because Atn(1) = 45³ = Pi/4
Const SamplesPerSec = 44100                             ' 44KHz
Const Amplitude = 2 ^ 15 - 1                            ' 16 BIT SAMPLES
Const Pause = 0
'--
Dim EventKey                As Byte                     ' 0 TO 128
Dim Duration                As Single                   ' 0 TO 1.00
'Dim Envelope(1 To 100)     As Single                   ' FOR 100 POINT ENVELOPE
Dim Envelope(1 To 30)       As Single                   ' FOR  30 POINT ENVELOPE
Dim OneWaveBuffer()         As Integer
Dim OutFile                 As Integer
Dim Frequency               As Long
Dim MidiKeyFreq(0 To 128)   As Long                     ' FREQUENCY FOR EACH MIDI NOTE

Public Sub Render(ByVal WaveFileName As String, _
                  ByVal EnvelopeFilename As String, _
                  ByVal TypeOfSample As Integer)
       
    'AppLog "RENDERING " & NoteCount & " NOTES TO FILE"
    '--
    If FileExist(WaveFileName) Then Kill WaveFileName   ' DELETE OLD FILE
    OpenWaveFile WaveFileName                           ' Example: SongfileName.wav
                 Seek #OutFile, 45                      ' DATA START
    LoadEnvelope EnvelopeFilename                       ' Example: Instrument.env
    LoadFrequencyDefinitions
    '--
    NoteCount = GetEventCount
    For i = 1 To NoteCount
        '--
        EventKey = Song.Events(i).MidiKey
        Duration = Song.Events(i).Duration * 4
        'AppLog "Event " & i & ")" & EventKey & "," & Duration
        '--
        If EventKey = Pause Then
        RenderNote 0, 1, TypeOfSample
        Else
        RenderNote EventKey, Duration, TypeOfSample
        End If
        '--
    Next
    
    CloseWaveFile
    WriteWaveHeader WaveFileName

End Sub

Private Function LoadEnvelope(ByVal FileName As String)

'   LOAD ENVELOPE FROM FILE TO MEMORY
    EnvFile = FreeFile
    Open FileName For Input As #EnvFile
    For i = 1 To UBound(Envelope())
    Input #EnvFile, Envelope(i)
    Next
    Close #EnvFile

End Function

Private Function LoadFrequencyDefinitions()

    FileName = App.Path & "\Data\MidiKey.Freq"
    Freqfile = FreeFile
    Open FileName For Input As Freqfile
    
    Do While Not EOF(Freqfile)
        'FORMAT IS: MIDIKEY,FREQUENCY,COMMENT
        Input #Freqfile, MidiKey
        Input #Freqfile, MidiKeyFreq(MidiKey)
        Input #Freqfile, Dummy 'Comment
        Loop
    
    Close #Freqfile, Comment
    
End Function

Private Function OpenWaveFile(FileName)

    OutFile = FreeFile
    Open FileName For Binary As #OutFile

End Function

Private Function CloseWaveFile()
    
    Close #OutFile

End Function

Private Function RenderNote(ByVal Note As Long, ByVal Duration As Single, ByVal TypeOfSample As Integer)
    
'   Note Should Be > 0 and < 128
'   If Note < 0 Then AppLog "Converting Note=0": Note = 0
'   If Note > 128 Then AppLog "Converting Note=128": Note = 128
        
    If Note = 0 Then 'PAUSE
       For Counter = 1 To (Frequency * 0.25)
       BufferToFile 0
       Next Counter
       Exit Function
       End If
       
'   CALCULATE THE FREQUENCY OF THE NOTE
    Frequency = MidiKeyFreq(Note)   ' MidiKey -> Frequency
'   If Frequency = 0 Then MsgBox "0 Frequency Error": Exit Function
        
'   NOW CREATE A WAVE SAMPLE OF THE CALCULATED FREQUENCY
    SampleCount = SamplesPerWave
    MakeSample SampleCount, TypeOfSample
    
'   NOT HAPPY WITH PART OF CODE BELOW ...QUICK AND DIRTY
'   ... WILL CHANGE LATER

    N = UBound(Envelope())
    Factor = 2.4
    TempoSpeed = Factor * N  ' RENDER SPACING HIGH VALUE = FASTER PLAY
    NumberOfWaves = Frequency / TempoSpeed
    
    ' BASED ON DURATION SELECT NUMBER OF POINTS
    N = Int(((N - 1) / 4) * Duration)
    For i = 1 To N
        For Counter = 1 To NumberOfWaves
        BufferToFile Envelope(i)
        Next Counter
    Next
    
    'Balance = Frequency - NumberOfWaves * 9
    'For Counter = 1 To Balance: BufferToFile2 Envelope(i): Next Counter

End Function

Private Function MakeSample(ByVal SampleCount As Long, ByVal TypeOfSample As SingleSample) As Long
    
    'AppLog "CREATING SINGLE WAVE SAMPLE WITH " & SampleCount & " SAMPLE POINTS"
    
'   CERTAIN EQUATIONS TO RENDER THE WAVE SAMPLES
'   WhiteNoise,SinePlusHarmonics,Sawtooth, SquareWave
'   CREDIT : Ulli 's Soundmaker Example. txtCodeId=64845
            
            
    ReDim OneWaveBuffer(SampleCount) As Integer
    StepValue = 2 * Pi / SampleCount
    
    Select Case TypeOfSample
    Case PureSine: ' PURE SINE WAVE
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Int(Sin(theta) * Amplitude)
         Next
    Case WhiteNoise
         For X = 0 To SampleCount - 1
         OneWaveBuffer(X) = Int((Rnd - Rnd) * Amplitude)
         Next
    Case SinePlusNoise ' PRODUCES SQUARISH DISTORTED WAVE
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Int((Sin(theta) + Rnd * 0.5)) * Amplitude
         'OneWaveBuffer(X) = Int((Sin(theta) * Rnd * 0.9)) * Amplitude
         AppLog OneWaveBuffer(X)
         Next
    Case SinePlusHarmonics
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         ' USE ODD EXPONENT TO PRESERVE SIGN OF SAMPLES
         OneWaveBuffer(X) = Int((Sin(theta) ^ 3) * Amplitude)
         Next
    Case Sawtooth
         A = Amplitude: S = SampleCount
         For X = 0 To SampleCount - 1
         OneWaveBuffer(X) = A / S * 2 * (X Mod S) - A
         Next
    Case SquareWave
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Amplitude * Sgn(Sin(theta))
         Next
    Case Other1  ' Sharp
         Step10 = StepValue * 10    ' Step3 also good
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         fb = (Sin(theta + Step10)) * Amplitude
         OneWaveBuffer(X) = Int((Sin(theta + fb)) * Amplitude)
         Next
    Case Other2  ' Vibratory
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Int((Sin(theta + Cos(theta / 2))) * Amplitude)
         Next
    Case Other3
         M = StepValue * 3
         N = StepValue * 5
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         F = (Sin(theta + M)) * Amplitude
         F = (Sin(theta + N) - F) * Amplitude
         OneWaveBuffer(X) = Int((Sin(theta + F)) * Amplitude)
         Next
    Case Other4
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Int(Sin(theta + Cos(theta)) * Amplitude)
         Next
    Case Other5
         For X = 0 To SampleCount - 1
         theta = theta + StepValue
         OneWaveBuffer(X) = Int((Sin(theta + Cos(theta / 0.5))) * Amplitude)
         Next
    End Select
    
End Function

Private Function BufferToFile(ByVal AmplitudeFactor As Single)

'   APPLY ENVELOPE TO ONE WAVE AND THEN WRITE IT TO WAVE FILE
'   AppLog "WRITING BUFFER TO FILE WITH AMPLITUDE FACTOR " & AmplitudeFactor

    First = LBound(OneWaveBuffer())
    Last = UBound(OneWaveBuffer()) - 1
    
    AmplitudeFactor = Abs(AmplitudeFactor)              ' \__ JUST BEING SURE
    If AmplitudeFactor > 1 Then AmplitudeFactor = 1     ' /
    
    For X = First To Last
    Put #OutFile, , CInt(OneWaveBuffer(X) * AmplitudeFactor)
    Next
   
End Function

Public Sub WriteWaveHeader(FileName)

'   WE WRITE THE HEADER AFTER ALL SAMPLES ARE RENDERED TO FILE
'   B'COZ WE DONT HAVE THE FILESIZE/DATASIZE INITIALLY
    
    WaveFile = FreeFile
    Open FileName For Binary As #WaveFile
    '-------------------------------------------'
    'SOME SETTING/CALCULATION BEFORE WE WRITE   '
    FileSize = LOF(WaveFile)                    '
    RIFFLength = FileSize - 8                   '
    DataLength = FileSize - 44                  '
    SampleRate = 44100                          ' KHz: 44100,22050,11025
    BitsPerSample = 16                          '
    Channels = 1                                ' = MONO
    BlockAlign = 2                              ' = BitsPerSample * Channels / 8
    BytesPerSecond = 88200                      ' = SampleRate * BitsPerSample * Channels / 8
    '-------------------------------------------'
    'RIFF HEADER FIRST                          '
    Put WaveFile, 1, "RIFF"                     '04  BYTES COMPLETED
    Put WaveFile, 5, CLng(RIFFLength)           '08  BYTES COMPLETED
    'THEN WAVE HEADER                           '
    Put WaveFile, 9, "WAVE"                     '12  Format
    Put WaveFile, 13, "fmt "                    '16  Subchunk1ID
    Put WaveFile, 17, CLng(16)                  '20  Subchunk1Size (16 for PCM)
    Put WaveFile, 21, CInt(1)                   '22  AudioFormat   (1 for PCM)
    Put WaveFile, 23, CInt(Channels)            '24  NumChannels Mono/Stereo
    Put WaveFile, 25, CLng(SampleRate)          '28  SamplesPerSecond
    Put WaveFile, 29, CLng(BytesPerSecond)      '
    Put WaveFile, 33, CInt(BlockAlign)          '
    Put WaveFile, 35, CInt(BitsPerSample)       '36
    Put WaveFile, 37, "data"                    '40  Subchunk2ID
    Put WaveFile, 41, CLng(DataLength)          '44  Subchunk2Size
    
    'ENTIRE HEADER WRITTEN                      '= 44 BYTES TOTAL
    'ALL BYTES FROM BYTE 45 IS DATA             '
    
    Close WaveFile
    
End Sub

Private Function SamplesPerWave() As Long

    SamplesPerWave = Int(SamplesPerSec / Frequency)
    
End Function

