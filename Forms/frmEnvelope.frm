VERSION 5.00
Begin VB.Form frmEnvelope 
   Caption         =   "Music Synthesis [Envelope Editor]"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin Synthesis.SeqEditor EnvelopeEdit 
      Height          =   3855
      Left            =   2280
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   4440
      Width           =   4695
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   0
      Pattern         =   "*.env"
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmEnvelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LastPoint = 100
Dim Envelope(1 To 100)           As Single       ' We use 100 point envelope
Dim EnvFile As Integer

Private Sub EnvelopeEdit_MouseMove(Button As Integer, Column As Integer, Row As Integer, CellDifference As Single)
    If Button = vbLeftButton Then
    EnvelopeEdit.DrawNote Column, Row, 0.25
    Envelope(Column) = Row
    End If
End Sub

Private Sub Form_Load()

    EnvelopeEdit.Columns = 100
    EnvelopeEdit.Rows = 100
    
End Sub

Sub LoadEnvelope(FileName)

    EnvFile = FreeFile
    Open FileName For Input As EnvFile
    For i = 1 To LastPoint
    Input #EnvFile, Envelope(i)
    Envelope(i) = 100 - Envelope(i) * 100
    Next
    Close

End Sub

Sub SaveEnvelope(FileName)

    EnvFile = FreeFile
    FileName = App.Path & "\Envelopes\Experiment.env"
    Open FileName For Output As EnvFile
    For i = 1 To LastPoint
    Print #EnvFile, 1 - Envelope(i) / 100
    Next
    Close
    
End Sub

Private Sub btnSave_Click()

    SaveEnvelope App.Path & "\Envelopes\Experiment.env"

End Sub

Private Sub btnLoad_Click()

    LoadEnvelope App.Path & "\Envelopes\Experiment.env"
    RefreshEditor

End Sub

Function RefreshEditor()

    For i = 1 To LastPoint
    EnvelopeEdit.DrawNote CInt(i), CInt(Envelope(i)), 0.25
    Next

End Function

