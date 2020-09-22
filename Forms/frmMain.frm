VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Music Synthesis"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   -4725
   ClientWidth     =   6105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSplash 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   3015
      Begin VB.Image Splash 
         Height          =   255
         Left            =   120
         Picture         =   "frmMain.frx":08A6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraRender 
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton btnRender 
         BackColor       =   &H00808000&
         Caption         =   "Render and Play"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00808000&
         TabIndex        =   8
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton btnStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   4680
         Width           =   735
      End
      Begin VB.ListBox SampleList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1620
         ItemData        =   "frmMain.frx":3B88A
         Left            =   120
         List            =   "frmMain.frx":3B8AF
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.FileListBox lstEnvelope 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2040
         Left            =   120
         Pattern         =   "*.env"
         TabIndex        =   5
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin MCI.MMControl MMC 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame fraEditor 
      BackColor       =   &H00000000&
      Height          =   4695
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.HScrollBar HScroll 
         Height          =   375
         LargeChange     =   20
         Left            =   960
         Max             =   180
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton btnZoom 
         Caption         =   "Zoom"
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   1320
         Width           =   855
      End
      Begin Synthesis.SeqEditor SeqEditor1 
         Height          =   975
         Left            =   600
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1720
      End
      Begin VB.Image ImgKeyboard 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   120
         Picture         =   "frmMain.frx":3B985
         Stretch         =   -1  'True
         Top             =   240
         Width           =   525
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3960
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileDivider1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export "
         Begin VB.Menu mnuExport2Midi 
            Caption         =   "As &Midi (SingleTrack)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuExport2Wave 
            Caption         =   "As &Wave (16Bit PCM)"
         End
         Begin VB.Menu mnuExport2Text 
            Caption         =   "As &Text"
         End
      End
      Begin VB.Menu mnuFileDivider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSampleType 
         Caption         =   "Sample Type"
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "PureSine"
            Index           =   0
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "SquareWave"
            Index           =   1
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "WhiteNoise"
            Index           =   2
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Sawtooth"
            Index           =   3
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "SinePlusNoise"
            Index           =   4
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "SinePlusHarmonics"
            Index           =   5
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Other1 (Sharp)"
            Index           =   6
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Other2 (Vibratory)"
            Index           =   7
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Other3"
            Index           =   8
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Other4"
            Index           =   9
         End
         Begin VB.Menu mnuSampleChoice 
            Caption         =   "Other5"
            Index           =   10
         End
      End
      Begin VB.Menu mnuEditorColors 
         Caption         =   "Editor Colors"
         Begin VB.Menu mnuOptionColors 
            Caption         =   "Plain &White"
            Index           =   0
         End
         Begin VB.Menu mnuOptionColors 
            Caption         =   "&Black Night"
            Index           =   1
         End
         Begin VB.Menu mnuOptionColors 
            Caption         =   "Chocolate &Brown"
            Index           =   2
         End
         Begin VB.Menu mnuOptionColors 
            Caption         =   "Leaf &Green"
            Index           =   3
         End
      End
      Begin VB.Menu mnuAppColors 
         Caption         =   "App Colors"
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Black"
            Index           =   0
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Pale Dull Blue"
            Index           =   1
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Ink Blue"
            Index           =   2
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Brown"
            Index           =   3
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Green"
            Index           =   4
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Light Brown"
            Index           =   5
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "Violet"
            Index           =   6
         End
         Begin VB.Menu mnuAppColorsOption 
            Caption         =   "DesignTime Setting"
            Index           =   99
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuLoadDemo 
         Caption         =   "Load &Demo Song"
      End
      Begin VB.Menu mnuShowDemo 
         Caption         =   "Show Demo ..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEnvelopeEditor 
         Caption         =   "Envelope Editor"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpDivider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' WE ARE USING FOLLOWING CONTROLS
'
' 1) A MULTIMEDIA CONTOL ON THE FORM O PLAY WAVE FILES
' 2) COMMONDIALOG FOR FILE/SAVE
' 2) EDITORCONTOL FOR EDITING OF KEYBOARD EVENTS

Const FullNote = 1#
Const ThreeFourthNote = 0.75
Const HalfNote = 0.5
Const QuarterNote = 0.25

Dim VisibleRows                 As Integer
Dim TotalEvents                 As Integer
Dim VisibleEvents               As Integer
Dim Column                      As Integer
Dim Row                         As Integer
Dim Period                      As Single
Dim ZoomFactor                  As Integer
Dim BaseMidiKey                 As Integer
Dim TypeOfSample                As Integer
Dim WaveFile                    As String
Dim FileNameOnly                As String

Private Sub Form_Load()
    
    SongFile.Saved = True   '\____ MUST BE DONE BEFORE BELOW INITIALISE
    mnuFileNew_Click        '/
    
    TotalEvents = Song.EventCount
    BaseMidiKey = Song.BaseMidiKey
    ZoomFactor = 10
    VisibleEvents = 30
    VisibleRows = 48
        
    SeqEditor1.ColorChoice = bgDarkBlue
    SeqEditor1.Columns = VisibleEvents
    SeqEditor1.Rows = VisibleRows
    SeqEditor1.CursorOn = True          ' PROPERTY BAG CODE IS STILL
    SeqEditor1.NumbersOn = True         ' NOT WRITTEN...SO HAVE TO DO IT RUNTIME
    
    HScroll.LargeChange = VisibleEvents
    HScroll.Max = TotalEvents - VisibleEvents
    CD.InitDir = App.Path & "\Songs" ' OR GET LASTUSE FROM REGISTRY
    lstEnvelope.Path = App.Path & "\Envelopes" ' This is a file listbox
    lstEnvelope.ListIndex = 0
    SampleList.ListIndex = 0
    SampleList_Click
    mnuAppColorsOption_Click 0
   
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next ' GoTo ResizeError
   
    MinHeight = 480 * 15: MinWidth = 640 * 15
    If Height < MinHeight Then Height = MinHeight
    If Width < MinWidth Then Width = MinWidth
    
    Gap = 90
    
'   LockWindowUpdate API
'   Disables/Enables drawing. Only one window can be locked at a time.
    
    LockWindowUpdate Me.hWnd
    
    ' RESIZE THE FRAMES
    fraRender.Height = ScaleHeight * 0.75
    fraEditor.Top = fraRender.Top
    fraEditor.Height = fraRender.Height
    FraSplash.Top = fraRender.Height + fraRender.Top
    FraSplash.Height = ScaleHeight - FraSplash.Top - Gap
    fraEditor.Width = ScaleWidth - fraEditor.Left - Gap * 2
    FraSplash.Width = ScaleWidth - FraSplash.Left - Gap * 2
     
    ' CONTROLS ON FIRST FRAME
    'LockWindowUpdate fraRender.hWnd
    'fraRender.Visible = False
    btnRender.Top = fraRender.Height - btnRender.Height - Gap
    btnStop.Top = btnRender.Top
    lstEnvelope.Height = btnRender.Top - lstEnvelope.Top - Gap
    'fraRender.Visible = True
    'LockWindowUpdate 0
    
    ' CONTROLS ON SECOND FRAME
    'LockWindowUpdate fraEditor.hWnd
    btnZoom.Left = ImgKeyboard.Left
    btnZoom.Top = fraEditor.Height - btnZoom.Height - Gap
    HScroll.Top = btnZoom.Top
    HScroll.Left = btnZoom.Left + btnZoom.Width
    HScroll.Width = fraEditor.Width - HScroll.Left - Gap * 2
    ImgKeyboard.Height = btnZoom.Top - ImgKeyboard.Top - Gap * 2
    'SeqEditor1.Visible = False
    SeqEditor1.Top = ImgKeyboard.Top
    SeqEditor1.Height = ImgKeyboard.Height
    SeqEditor1.Left = ImgKeyboard.Left + ImgKeyboard.Width
    SeqEditor1.Width = fraEditor.Width - SeqEditor1.Left - Gap * 2
    'SeqEditor1.Visible = True
    'LockWindowUpdate 0
    
    ' CONTROLS ON THIRD FRAME
    'LockWindowUpdate FraSplash.hWnd
    Splash.Visible = False
    Splash.Left = Gap * 2
    Splash.Height = FraSplash.Height - Splash.Top - Gap * 2
    Splash.Width = FraSplash.Width - Gap * 6
    Splash.Visible = True
    'LockWindowUpdate 0
    
ResizeError:
'   RESIZE ERROR IS NOT CRITICAL IN MOST CASES. WE JUST CONTINUE
    LockWindowUpdate 0
    RefreshEditor
    Move 0, 0

End Sub

Private Sub mnuAbout_Click()
' TODO : WILL WRITE A PROPER ABOUT...LATER
    Message = About("NAME") & " by " & About("Author") _
              & vbCrLf & "Version " & About("VERSION")

              
    MsgBox Message, , About("title")
End Sub

Private Sub mnuAppColorsOption_Click(Index As Integer)
    
    
    
    Select Case Index
    Case 0: ColorChoice = 0             ' Black
    Case 1: ColorChoice = &H80000011    ' Pale DullBlue
    Case 2: ColorChoice = &HC00000      ' InkBlue '&H808000 'Blue
    Case 3: ColorChoice = &H404080      ' Brown
    Case 4: ColorChoice = &H9000&       ' Green
    Case 5: ColorChoice = &H8080&       ' LightBrown
    Case 6: ColorChoice = &HAA7070      ' Violet
    Case 99: ColorChoice = BackColor    ' For Experiment Only
    End Select
    
    BackColor = ColorChoice
    fraRender.BackColor = ColorChoice
    fraEditor.BackColor = ColorChoice
    FraSplash.BackColor = ColorChoice
    
    Static OldChoice
    mnuAppColorsOption(OldChoice).Checked = False
    OldChoice = Index
    mnuAppColorsOption(OldChoice).Checked = True
    
End Sub


'=========================================================
' CODING FOR THE EDITOR. EASY :-)
'=========================================================
Private Sub SeqEditor1_LeftClick(Column As Integer, Row As Integer, CellDifference As Single)

    SeqEditor1.DrawNote Column, Row, FullNote
    Song.Events(Column).MidiKey = GridMidiConvert(Row)
    Song.Events(Column).Duration = FullNote
    SongFile.Saved = False

End Sub

Private Sub SeqEditor1_RightClick(Column As Integer, Row As Integer, CellDifference As Single)

    SeqEditor1.EraseNote Column, Row
    Song.Events(Column).MidiKey = GridMidiConvert(0)
    Song.Events(Column).Duration = 0
    SongFile.Saved = False

End Sub

Private Sub SeqEditor1_MouseMove(Button As Integer, Column As Integer, Row As Integer, CellDifference As Single)

    If (Button = vbLeftButton) And (SeqEditor1.MouseDownCol = Column) Then
        '--
        Select Case CellDifference
        Case Is < 0.3:      DELTA = QuarterNote
        Case 0.3 To 0.6:    DELTA = HalfNote
        Case 0.6 To 0.8:    DELTA = ThreeFourthNote
        Case Else:          DELTA = FullNote
        End Select
        '--
        SeqEditor1.DrawNote Column, Row, CSng(DELTA)
        Song.Events(Column).MidiKey = GridMidiConvert(Row)
        Song.Events(Column).Duration = DELTA
        SongFile.Saved = False
    End If

End Sub

'=========================================================
' MENU ROUTINES FOLLW:
' ROUTINES TO OPEN,SAVE FILES, ETC
'=========================================================

Private Sub mnuFileNew_Click()
    
    VerifySaveCurrentSong
    CreateNewSong
    RefreshEditor
    
    FileNameOnly = "Untitled.sng" ' Dir$(SongFile.Name)
    Caption = About("NAME") & " [" & FileNameOnly & "]"
    HScroll.Value = HScroll.Min
    HScroll_Change
    
End Sub

Private Sub mnuFileOpen_Click()
    
    VerifySaveCurrentSong
    LoadSong GetOpenFileName()
    RefreshEditor
    
    FileNameOnly = Dir$(SongFile.Name)
    Caption = About("NAME") & " [" & FileNameOnly & "]"
    HScroll.Value = HScroll.Min
    HScroll_Change

End Sub


Private Sub mnuFileSave_Click()

Rem TODO:IF FILE ALREADY EXIST...ASK FOR OVERWRITE PERMISSION

    If SongFile.Saved Then Exit Sub ' NO NEED TO SAVE
    
    'If Dir$(SongFile.Name) = "" Then ' IF STILL NOT SAVED
    
    ' IF WE DONT HAVE A NAME GET IT
    If Right(SongFile.Name, 12) = "Untitled.sng" Then
       FileName = GetSaveFileName()
       If FileName = "" Then Exit Sub
       SongFile.Name = FileName
       End If
    SaveSong SongFile.Name

End Sub

Private Sub mnuFileSaveAs_Click()

Rem TODO:IF FILE ALREADY EXIST...ASK FOR OVERWRITE PERMISSION

    FileName = GetSaveFileName()
    If FileName = "" Then Exit Sub
    SongFile.Name = FileName
    SaveSong SongFile.Name
    
    FileNameOnly = Dir$(SongFile.Name)
    Caption = About("NAME") & " [" & FileNameOnly & "]"
    
 
End Sub

Private Sub mnuOptionColors_Click(Index As Integer)
    SeqEditor1.ColorChoice = Index
    RefreshEditor
End Sub
    
Private Sub mnuExport2Wave_Click()
    Export2Wave
End Sub

Private Sub mnuExport2Text_Click()

Rem THIS IS A QUICK EXPORT OF THE SONG TO A TEXT FILE _
    TODO:REWRITE TO GIVE USER A CHOICE OF EXPORT FILENAME

'   LOCATE LAST NON-ZERO EVENT
    N = GetEventCount
    If N = 0 Then
        Message = "There are no events to Export. Cannot Export"
        MsgBox Message, vbExclamation, "Yups !!"
        Exit Sub
        End If

'   WRITE HEADER
    StrLine = String(80, "=")
    txtFile = FreeFile
    FileName = SongFile.Name & "-export.txt"
    Open SongFile.Name & "-export.txt" For Output As #txtFile
    Print #txtFile, StrLine
    Print #txtFile, "Exported From File:"; SongFile.Name
    Print #txtFile, "Date:"; Format$(Date$, "DD MMM,YYYY")
    Print #txtFile, "Event Count:"; N
    Print #txtFile, StrLine
    Print #txtFile,
    Print #txtFile, "Event#"; Tab(10); "MidiKey#"; Tab(25); "Duration"
    Print #txtFile, "-----"; Tab(10); "---------"; Tab(25); "--------"
    
'   EXPORT TO TEXT FILE
    For i = 1 To N
    Print #txtFile, i; _
                    Tab(12); Song.Events(i).MidiKey; _
                    Tab(27); Song.Events(i).Duration
    Next
    Close #txtFile
    Message = "The song has been exported to" & vbCrLf & FileName
    MsgBox Message, vbInformation, "Export"
    
End Sub

Private Sub mnuSampleChoice_Click(Index As Integer)
    
    mnuSampleChoice(TypeOfSample).Checked = False
    TypeOfSample = Index
    mnuSampleChoice(TypeOfSample).Checked = True
    SampleList.ListIndex = Index
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuLoadDemo_Click()

    FileName = (App.Path & "\Songs\Demo.sng")
    LoadSong FileName
    RefreshEditor
    
    FileNameOnly = Dir$(SongFile.Name)
    Caption = About("NAME") & " [" & FileNameOnly & "]"
    
    HScroll.Value = HScroll.Min
    HScroll_Change
    
End Sub


'=========================================================
' ROUTINES TO PLAY THE WAVE FILE
'=========================================================

Sub WavFileLoad(ByVal FileName As String)

Rem BEFORE WE 'OPEN' AND 'PLAY' ANOTHER WAVE FILE, WE MUST _
    STOP ANY OPEN FILE THAT IS ALREADY PLAYING. OTHERWISE _
    THE SECOND FILE WONT PLAY
    
    If Dir(FileName) = "" Then
            Message = "Error Loading Rendered Wav File. Can't Play"
            MsgBox Message, , , "Error"
    Else:   MMC.DeviceType = "WAVEAUDIO"
            MMC.Notify = True
            MMC.FileName = FileName
            WavFileStop
            End If

End Sub

Sub WavFilePlay()

    If WaveFile = "" Then
        Message = "Please select a file from the list to be played"
        MsgBox Message, vbExclamation, "Yups !!"
        Exit Sub
        End If

    WavFileLoad WaveFile
    SetButtons False, False 'Play=0,Stop=0
    If MMC.Mode = mciModeNotOpen Then MMC.Command = "open"
    If MMC.Mode = mciModePause Then
            MMC.Command = "pause"
    Else:   MMC.Command = "play"
            End If

    Do: DoEvents
    Loop Until MMC.Mode = mciModePlay
    SetButtons False, True  'Play=0,Stop=1

End Sub

Sub WavFilePause()

    If MMC.Mode <> mciModePlay Then Exit Sub
    SetButtons False, False  'Play=0,Stop=0
    MMC.Command = "pause"
    Do: DoEvents
    Loop Until MMC.Mode = mciModePause
    SetButtons True, False   'Play=1,Stop=0
  
End Sub
Sub WavFileStop()

    SetButtons False, False   'Play=0,Stop=0
    MMC.Command = "stop"
    MMC.Command = "close"
    Do: DoEvents
    Loop Until MMC.Mode = mciModeNotOpen
    SetButtons True, False    'Play=1,Stop=0
    
End Sub

Private Sub MMC_Done(NotifyCode As Integer)
    WavFileStop
End Sub

Sub SetButtons(valPlay As Boolean, valStop As Boolean)
    btnRender.Enabled = valPlay
    btnStop.Enabled = valStop
End Sub

'=========================================================
' CODING FOR VARIOUS OTHER CONTROLS FOLLOW
'=========================================================

Private Sub btnZoom_Click()
    ZoomFactor = -ZoomFactor ' TOGGLE ZoomFactor
    VisibleEvents = VisibleEvents + ZoomFactor
    SeqEditor1.Columns = VisibleEvents
    HScroll.LargeChange = VisibleEvents
    HScroll.Max = TotalEvents - VisibleEvents
    RefreshEditor
End Sub

Private Sub HScroll_Change()
    SeqEditor1.LeftCol = HScroll.Value + 1
    RefreshEditor
End Sub

Private Sub SampleList_Click()
    mnuSampleChoice(TypeOfSample).Checked = False
    TypeOfSample = SampleList.ListIndex
    mnuSampleChoice(TypeOfSample).Checked = True
End Sub

Private Sub btnRender_Click()
    If Export2Wave Then WavFilePlay
End Sub

Private Sub btnStop_Click()
    WavFileStop
End Sub

'=========================================================
' VARIOUS OTHER HELPER ROUTINES HERE
'=========================================================
Sub RefreshEditor()

    For N = 0 To VisibleEvents - 1
    Column = CInt(N + SeqEditor1.LeftCol)
    Row = GridMidiConvert(Song.Events(Column).MidiKey)
    Period = Song.Events(Column).Duration
    SeqEditor1.DrawNote Column, Row, Period
    Next

End Sub

Function GridMidiConvert(ByVal Row As Integer) As Integer

Rem THIS FUNCTION DOES BOTH INTERCONVERSIONS. _
    SO WE DONT NEED THE FOLLOWING SEPARATE FUNCTIONS
    '--
    'Function Midi2Grid(Row As Integer) As Integer ' same as Grid2Midi
    '    Midi2Grid = BaseMidiKey + VisibleRows - Row
    'End Function
    '--
    'Function Grid2Midi(Row As Integer) As Integer ' same as Midi2Grid
    '    Grid2Midi = BaseMidiKey + VisibleRows - Row
    'End Function
    '--
    
    Select Case Row
    Case 0:    GridMidiConvert = 0
    Case Else: GridMidiConvert = BaseMidiKey + VisibleRows - Row
    End Select
    
End Function

Sub VerifySaveCurrentSong()

Rem CANNOT USE FOLLOWING -- VB CALCULATES BOTH
    ' SaveIt = (Not (SongFile.Saved)) And UserWants2Save
    ' If Not (SaveIt) Then Exit Sub
    
    If SongFile.Saved Then Exit Sub
    If Not (UserWants2Save) Then Exit Sub
    If SongFile.Name = "" Then SongFile.Name = GetSaveFileName()
    SaveSong SongFile.Name
    
End Sub

Function UserWants2Save() As Boolean
    Prompt = "Do you want to save the changes to current file?"
    Answer = MsgBox(Prompt, vbQuestion + vbYesNo)
    UserWants2Save = (Answer = vbYes)
End Function

Function GetOpenFileName() As String
    On Error GoTo ErrorHandler:
    
    CD.DialogTitle = "Open"
    CD.CancelError = False
    CD.Filter = "Song File (*.sng)|*.sng"
    CD.Flags = cdlOFNExplorer + cdlOFNFileMustExist
    CD.ShowOpen
    
    If Len(CD.FileName) = 0 Then
            GetOpenFileName = ""
    Else:   GetOpenFileName = CD.FileName
            CD.InitDir = ""
            End If
      
ErrorHandler:

End Function

Function GetSaveFileName() As String

    On Error GoTo ErrorHandler

    CD.DialogTitle = "Save"
    CD.CancelError = False      'CD.CancelError = True
    CD.Filter = "Song File (*.sng)|*.sng"
    CD.DefaultExt = "*.sng"
    CD.Flags = cdlOFNHideReadOnly
    CD.ShowSave
    
    If Len(CD.FileName) = 0 Then
            GetSaveFileName = ""
    Else:   GetSaveFileName = CD.FileName
            CD.InitDir = ""
            End If
    
ErrorHandler:

End Function
 
Function Export2Wave() As Boolean

    NoteCount = GetEventCount
    If NoteCount = 0 Then
        Message = "There is nothing to render." _
                   & vbCrLf & "Please open a *.sng file or" _
                   & vbCrLf & "create some events on the " _
                   & vbCrLf & "editor for rendering."
        MsgBox Message, vbExclamation, "Yups !!"
        Export2Wave = False
        Exit Function
    End If
    '---------------
    'FileNameOnly = Mid$(SongFile.Name, InStrRev(SongFile.Name, "\") + 1) ' TEMP STRING
    FileNameOnly = Dir$(SongFile.Name)
    If FileNameOnly = "" Then FileNameOnly = "Untitled.sng"
    WaveFile = App.Path & "\Renders\" & _
               Replace(FileNameOnly, ".sng", ".wav")
    EnvelopeFile = App.Path & "\Envelopes\" & lstEnvelope.FileName
    '---------------
    MousePointer = 11
    btnRender.Enabled = False   ' DONT WANT USER TO DISTURB US
    Synt.Render WaveFile, EnvelopeFile, TypeOfSample
    Export2Wave = True
    'Message = "The file has been rendered." _
    '          & vbCrLf & "Press the Play Button to " _
    '          & vbCrLf & "play the rendered file."
    'MsgBox Message, vbInformation, "Render Complete"
    btnRender.Enabled = True
    MousePointer = 0
    
      
End Function


