Attribute VB_Name = "modSongFileRoutines"

Sub LoadSong(FileName)

    If FileName = "" Then Exit Sub
    On Error GoTo OpenError

    Handle = FreeFile
    Open FileName For Binary As Handle
    Get Handle, , Song
    Close Handle
        
    SongFile.Name = FileName        ' MUST BE LAST LINE

OpenError:

End Sub

Sub SaveSong(FileName)

    On Error GoTo SaveError
    Handle = FreeFile
    Open FileName For Binary As Handle
    Put Handle, , Song
    Close Handle
    SongFile.Saved = True
       
SaveError:

End Sub

Public Sub CreateNewSong()

    Dim NewSong As typSong
    
    NewSong.Author = ""
    NewSong.Comments = "Created by " & About("NAME")
    NewSong.Date = Date
    NewSong.FileID = FileID
    NewSong.DataID = DataID
    NewSong.MajorVersion = MajorVersion
    NewSong.MinorVersion = MinorVersion
    NewSong.EventCount = 200
    NewSong.BaseMidiKey = 48
    
    Let Song = NewSong
    SongFile.Name = App.Path & "\Songs\Untitled.sng"
    SongFile.Saved = False
     
End Sub

Public Function About(Info As String)

    Select Case UCase(Info)
    Case "NAME":        About = AppName
    Case "VERSION":     About = MajorVersion & "." & MinorVersion
    Case "TITLE":       About = AppTitle
    Case "AUTHOR":      About = AppAuthor
    End Select
    
End Function

Public Function GetEventCount() As Integer

'   LOCATE LAST NON-ZERO EVENT
    For i = Song.EventCount To 1 Step -1
        A = (Song.Events(i).MidiKey = 0)
        If Not A Then GetEventCount = i: Exit Function
    Next
    GetEventCount = 0
             
End Function
