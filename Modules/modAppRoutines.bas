Attribute VB_Name = "modAppRoutines"
Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Global Const AppTitle = "Music Synthesis Example"
Global Const AppName = "Music Synthesis"
Global Const AppAuthor = "Akiti Yadav"
Global Const MajorVersion = 0
Global Const MinorVersion = 1
Global Const DataID = "KBEV"
Global Const FileID = "SNG1"

'----------------------------------------------------------------------
' SONG FILE DECLARES
'----------------------------------------------------------------------
Type typSongInfo
     Name               As String
     Saved                  As Boolean
End Type

Type typMidiEvent
     MidiKey                As Byte         ' 0 to 128
     Duration               As Single       '
     Vocal                  As Integer      ' Reserved
     End Type
     
Type typSong                        ' All the data from the song file
     FileID                 As String * 4   ' "SNG1" Simple Song File Type 1
     MajorVersion           As Byte         ' = 0 = Experimental
     MinorVersion           As Byte         ' = 0
     Comments               As String * 50  ' Any Comment
     Author                 As String * 20  ' Author of file
     BaseMidiKey            As Byte         ' Lowest Key MidiValue on Editor
     Reserved               As String * 9   ' 9 bytes for flags,etc
     Date                   As Date         ' Creation date
     DataID                 As String * 4   ' "KBEV" Keyboard Event
     EventCount             As Integer      ' =200 for SNG1 file
     Events(1 To 200)       As typMidiEvent ' We use 200 events
     End Type

Global Song                 As typSong
Global SongFile             As typSongInfo
Global Synt                 As New clsSynthMusic
Global Message              As String           ' FOR THOSE MESSAGE BOXES




'==================================================='

' VERIFY WHETHER FILE EXISTS
Public Function FileExist(ByVal FileName As String) As Boolean
    Select Case Len(Dir(FileName))
    Case 0: FileExist = False
    Case Else: FileExist = True
    End Select
End Function

Sub AppLog(S)
    Debug.Print S
End Sub


