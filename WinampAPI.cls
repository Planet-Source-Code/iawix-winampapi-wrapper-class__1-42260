VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinampAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ WinampAPI.cls - Included with BlueAmp Build 114 - www.iawix.com ++
'++               Written By Steve F. of iawix.com                  ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This code can be used in FREEWARE software only.  You can NEVER ++
'++ use this code/class or any portion of it in a piece of software ++
'++ that is supported by ads, sold commercially, installs third     ++
'++ party commercial software, uses spyware, or sends user          ++
'++ information back to a commercial server.                        ++
'++ Basically, try to make a cent off it, and I will kick your ass. ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This class is used in BlueAmp to reference Winamp 2's API in a  ++
'++ familiar object interface.  It was designed to make using       ++
'++ Winamp's API easy.  Customize it so it does that for you.  This ++
'++ version sports the latest in BlueAmp's technology.  It's lean   ++
'++ and fast.  Have fun!  Questions/Comments can be directed to     ++
'++ steve < AT > iawix.com                                          ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ And remember to visit IAWIX.COM and give me mad props please!   ++
'++ I'm always interested in how you use this code so drop me an    ++
'++ email and link me to your project!                              ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Example to start using Winamp's API:
'Public Winamp As New WinampAPI
'**It should work from then on if Winamp is open.

Option Explicit

Private hwndWinamp As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal WndID As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const WM_COMMAND = &H111
Private Const WM_USER = &H400

Private Enum WM_USER_MSGS
    wmVersion = 0
    wmStartPlayback = 100
    wmClearPlaylist = 101
    wmBeginPlayingTrack = 102
    wmPlaybackStatus = 104
    wmTrackTimeInfo = 105 'If data is 0, returns the position in milliseconds of playback. If data is 1, returns current track length in seconds. Returns -1 if not playing or if an error occurs.
    wmSeekTrack = 106 'Seeks within the current track. The offset is specified in 'data', in milliseconds.
    wmWritePlaylist = 120 'Writes out the current playlist to Winampdir\winamp.m3u, and returns the current position in the playlist.
    wmSetPlaylistPosition = 121 'Sets the playlist position to the position specified in tracks in 'data'.
    wmSetVolume = 122 'Sets the volume to 'data', which can be between 0 (silent) and 255 (maximum).
    wmSetPanning = 123 'Sets the panning to 'data', which can be between 0 (all left) and 255 (all right).
    wmPlaylistLength = 124 'Returns length of the current playlist, in tracks.
    wmPlaylistPosition = 125 'Returns the position in the current playlist, in tracks (requires Winamp 2.05+).
    wmTrackRateInfo = 126 'Retrieves info about the current playing track. Returns samplerate (i.e. 44100) if 'data' is set to 0, bitrate if 'data' is set to 1, and number of channels if 'data' is set to 2. (requires Winamp 2.05+)
    'I left out getting equalizer data.  I don't want it.
    wmBookmark = 129 'Adds the specified file to the Winamp bookmark list
    wmRestartWinamp = 135
End Enum

Private Enum WM_COMMAND_MSGS
    wmPrevTrack = 40044
    wmNextTrack = 40048
    wmPlayTrack = 40045
    wmPauseUnpause = 40046
    wmStopTrack = 40047
    wmFadeStop = 40147
    wmStopAfterCurrentTrack = 40157
    wmFastForward5 = 40148
    wmFastRewind5 = 40144
    wmGotoStartOfPlaylist = 40154
    wmGotoEndOfPlaylist = 40158
    wmOpenFileDlg = 40029
    wmOpenURLDlg = 40155
    wmFileInfoBox = 40188
    wmSetTimeElapsed = 40037
    wmSetTimeRemaining = 40038
    wmToggleAboutBox = 40041
    wmToggleTitleAutoScroll = 40189
    wmRaiseVolumePercent = 40058
    wmLowerVolumePercent = 40059
    wmToggleRepeat = 40022
    wmToggleShuffle = 40023
    wmJumpToFileDialog = 40194
    wmOpenSkinSelector = 40219
    wmCloseWinamp = 40001
    wmPlayAudioCD = 40323
End Enum

Private Sub Class_Initialize()
    Findhwnd
End Sub

Public Property Get WinampVersion() As String
    WinampVersion = SendUser(wmVersion)
End Property

Private Function SendUser(id As WM_USER_MSGS, Optional Data As Long) As Long
    Findhwnd 'Look for Winamp
    SendUser = SendMessage(hwndWinamp, WM_USER, Data, id)
End Function

Private Function SendCommand(id As WM_COMMAND_MSGS, Optional Data As Long) As Long
    Findhwnd 'Look for Winamp
    Dim lngTimer As Long
    lngTimer = Timer
    
    SendCommand = PostMessage(hwndWinamp, WM_COMMAND, id, Data)
    
    Do Until SendCommand <> 0
        If Timer = lngTimer + 10 Or Timer < 10 Then Exit Function
        DoEvents
    Loop

End Function

Public Property Get SongTitle() As String
    On Error Resume Next
    
    Findhwnd 'Look for Winamp
        
    Dim pstrTitle As String
    pstrTitle = String$(1024, Chr$(8))
    
    GetWindowText hwndWinamp, pstrTitle, 1024
    pstrTitle = Replace$(pstrTitle, Chr$(8), vbNullString)
    If (InStr(1, pstrTitle, " - Winamp")) > 0 Then
        pstrTitle = Left$(pstrTitle, InStr(1, pstrTitle, " - Winamp")) 'Get rid of trailing title.
    End If
    
    'Get rid of track
    If InStr(1, pstrTitle, ". ") > 0 Then pstrTitle = Trim$(Right$(pstrTitle, Len(pstrTitle) - InStr(1, pstrTitle, ". ")))
    
    SongTitle = pstrTitle
End Property

Public Property Get IsRunning() As Boolean
    Findhwnd
    
    Select Case hwndWinamp
        Case 0
            IsRunning = False
        Case Else
            IsRunning = True
    End Select
End Property

Private Sub Findhwnd()
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
End Sub

Public Sub Previous()
    SendCommand wmPrevTrack
End Sub

Public Sub Play()
    SendCommand wmPlayTrack
End Sub

Public Sub Pause()
    SendCommand wmPauseUnpause
End Sub

Public Sub Stop_()
    SendCommand wmStopTrack
End Sub

Public Sub Next_()
    SendCommand wmNextTrack
End Sub

Public Property Get TrackTime() As String
    Dim secIntoTrack As Integer, secTrack As Integer
    
    secIntoTrack = CInt((SendUser(wmTrackTimeInfo, 0)) / 1000)
    secTrack = SendUser(wmTrackTimeInfo, 1)
    
    TrackTime = Format$(Fix(secIntoTrack / 60), "#0") & ":" & Format$(Fix(secIntoTrack Mod 60), "00") & " / " & Format$(Fix(secTrack / 60), "#0") & ":" & Format$(Fix(secTrack Mod 60), "00")
End Property

Public Property Get TrackTimePercent() As Long
    TrackTimePercent = CLng(((SendUser(wmTrackTimeInfo, 0)) / 1000) / (SendUser(wmTrackTimeInfo, 1)) * 100)
End Property

Public Function OtherUser(ByVal id As Long, Optional ByVal Data As Long) As Long
'For lack of more time, I have simply used a generic function
'for methods I have not implemented.  Should you want to expand
'upon this DLL, this is what you'd get rid of and write your own
'custom functions (like above) to handle anything you'd call here.
    OtherUser = SendUser(id, Data)
End Function

Public Function OtherCommand(ByVal id As Long, Optional ByVal Data As Long) As Long
'For lack of more time, I have simply used a generic function
'for methods I have not implemented.  Should you want to expand
'upon this DLL, this is what you'd get rid of and write your own
'custom functions (like above) to handle anything you'd call here.
    OtherCommand = SendCommand(id, Data)
End Function

Public Property Get Status() As Long
    'Returns the status of playback.
    'If 'ret' is 1, Winamp is playing.
    'If 'ret' is 3, Winamp is paused.
    'Otherwise, playback is stopped.
    Status = SendUser(wmPlaybackStatus)
End Property

Public Sub EndWinamp()
    'Quits Winamp
    SendCommand wmCloseWinamp
End Sub

Public Sub VolumeUp()
    'Raises Volume 1 percent.
    SendCommand wmRaiseVolumePercent
End Sub

Public Sub VolumeDown()
    'Lowers Volume 1 percent.
    SendCommand wmLowerVolumePercent
End Sub

Public Property Get TrackNumber() As Integer
    On Error Resume Next
    
    Findhwnd 'Look for Winamp
        
    Dim pstrTitle As String
    pstrTitle = String$(1024, Chr$(8))
    
    GetWindowText hwndWinamp, pstrTitle, 1024
    pstrTitle = Replace$(pstrTitle, Chr$(8), vbNullString)
        
    'Get rid of track info
    pstrTitle = Trim$(Left$(pstrTitle, InStr(1, pstrTitle, ". ")))
    
    TrackNumber = CInt(pstrTitle)
End Property

Public Property Get TracksInPlaylist() As Integer
    'Returns number of tracks in playlist.
    TracksInPlaylist = CInt(SendUser(wmPlaylistLength))
End Property

Public Sub JumpToFile()
    'Opens jump to song dialog
    Dim tempHwnd As Long
            
    ShowWindow hwndWinamp, 9
    SetForegroundWindow hwndWinamp
    BringWindowToTop hwndWinamp
    'SendMessage hwndWinamp, &H1C, 0, 0
    SendCommand wmJumpToFileDialog
    tempHwnd = FindWindow("Jump to file", vbNullString)
    ShowWindow tempHwnd, 9
    SetForegroundWindow tempHwnd
    BringWindowToTop tempHwnd
End Sub

Public Sub ToggleShuffle()
    'Toggles shuffle button
    SendCommand wmToggleShuffle
End Sub

Public Sub FastForward()
    'Fast Forwards
    SendCommand wmFastForward5
End Sub

Public Sub Rewind()
    'Rewinds
    SendCommand wmFastRewind5
End Sub
