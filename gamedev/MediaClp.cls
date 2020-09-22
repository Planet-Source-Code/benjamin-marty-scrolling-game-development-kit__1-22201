VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "References a multimedia file that can be played (usually in response to an event).  Owned and indexed by MediaManager (= project)."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MediaClp.cls - Multimedia Clip Class
'
'======================================================================

Option Explicit

Public Enum MEDIA_FLAGS
    MEDIA_KEEP_LOADED = 1
    MEDIA_VIDEO = 2
    MEDIA_LOOP = 8
    MEDIA_MODAL = 16
    MEDIA_SUSPEND_OTHERS = 32
    MEDIA_WAITING_ON_SUSPEND = 64
    MEDIA_SUSPEND = 128
    'MEDIA_AUDIO = 256 This should go in version 1.3
End Enum

' Persistable Properties
Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this clip as indexed by the MediaManager object."
Public strMediaFile As String
Attribute strMediaFile.VB_VarProcData = ";Data"
Attribute strMediaFile.VB_VarDescription = "The path to the clip file as stored in the project file."
Public OutputX As Integer
Attribute OutputX.VB_VarProcData = ";Appearance"
Attribute OutputX.VB_VarDescription = "For video clips, set the left side of the output."
Public OutputY As Integer
Attribute OutputY.VB_VarProcData = ";Appearance"
Attribute OutputY.VB_VarDescription = "For video clips, set the top of the output."
Public Flags As Integer
Attribute Flags.VB_VarProcData = ";Behavior"
Attribute Flags.VB_VarDescription = "Determines various aspects of how the clip behaves."
Public Volume As Integer
Attribute Volume.VB_VarProcData = ";Appearance"
Attribute Volume.VB_VarDescription = "How loud this clip should be played (0=max)."

Dim strLoadedMedia As String
Dim strPlayingMedia As String
Dim strFadeToMedia As String
Dim nFadeVolume As Integer
Dim FGM As FilgraphManager
Public bCancelModalMedia As Boolean
Attribute bCancelModalMedia.VB_VarProcData = ";Behavior"
Attribute bCancelModalMedia.VB_VarDescription = "Set to True to stop a modal clip from playing (in response to an event)."

Public Sub FadeOut()
Attribute FadeOut.VB_Description = "Start this clip on a process of fading out to silence."
    strFadeToMedia = ""
End Sub

Public Sub Play(Optional ByVal strNewMedia As String = "")
Attribute Play.VB_Description = "Start or continue playing this clip, optionally fading it to a new media file."
    Dim bv As IBasicVideo
    Dim FInfo As IFilterInfo
    Dim FCol As IAMCollection
    Dim FIdx As Integer
    Dim VidWin As IVideoWindow
    Dim VidInf As IBasicVideo
    Dim AudInf As IBasicAudio
    Dim MPos As IMediaPosition
    Dim nState As Long
    Dim strErr
    
    On Error GoTo PlayErr

    If strNewMedia <> "" Then strFadeToMedia = strNewMedia

    If strFadeToMedia <> strPlayingMedia Or strPlayingMedia = "" Then
        If strPlayingMedia = "" Then
            If strFadeToMedia <> "" Then
                strPlayingMedia = strFadeToMedia
            Else
                strPlayingMedia = strMediaFile
                strFadeToMedia = strMediaFile
            End If
            If strLoadedMedia <> strPlayingMedia Then
                Set FGM = Nothing
            Else
                If (Flags And MEDIA_FLAGS.MEDIA_SUSPEND_OTHERS) <> 0 And (Flags And MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND) = 0 Then
                    Flags = Flags Or MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND
                    If strFadeToMedia = "" Then strFadeToMedia = strLoadedMedia
                    Exit Sub
                End If
            End If
        Else
            Set MPos = FGM
            If MPos.CurrentPosition < MPos.Duration And MPos.CurrentPosition > 0 Then
                If nFadeVolume > -10000 Then
                    nFadeVolume = Int((nFadeVolume - 2) * 1.1)
                    If nFadeVolume < -10000 Then nFadeVolume = -10000
                End If
                If nFadeVolume = -10000 Then
                    If Flags And 256 Then ' Replace 256 with MEDIA_FLAGS.MEDIA_AUDIO in version 1.3
                        Set AudInf = FGM
                        AudInf.Volume = Volume
                    End If
                    StopClip False
                    If strFadeToMedia <> "" Then Set FGM = Nothing
                    strPlayingMedia = strFadeToMedia
                End If
            Else
                StopClip True
                If strFadeToMedia <> "" Then Set FGM = Nothing
                strPlayingMedia = strFadeToMedia
            End If
        End If
    Else
        If nFadeVolume < Volume Then
            If nFadeVolume <= -5000 Then
                nFadeVolume = nFadeVolume + 500
            Else
                nFadeVolume = nFadeVolume + 100
            End If
            If nFadeVolume > Volume Then nFadeVolume = Volume
        Else
            nFadeVolume = Volume
        End If
    End If
    
    Flags = Flags And Not MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND
    
    If FGM Is Nothing And strPlayingMedia <> "" Then
        Set FGM = New FilgraphManager
        ChDir PathFromFile(Prj.ProjectPath)
        ChDrive Left$(PathFromFile(Prj.ProjectPath), 1)
        FGM.RenderFile strPlayingMedia
        strLoadedMedia = strPlayingMedia
        Set FCol = FGM.FilterCollection
        Flags = Flags And Not MEDIA_FLAGS.MEDIA_VIDEO
        Flags = Flags And Not 256 ' Replace 256 with MEDIA_FLAGS.MEDIA_AUDIO in version 1.3
        For FIdx = 0 To FCol.Count - 1
            FCol.Item FIdx, FInfo
            If InStr(FInfo.Name, "Video") > 0 Then
                Flags = Flags Or MEDIA_FLAGS.MEDIA_VIDEO
            End If
            If InStr(FInfo.Name, "Sound") > 0 Or InStr(FInfo.Name, "Audio") > 0 Or InStr(FInfo.Name, "MIDI") > 0 Then
                Flags = Flags Or 256 ' Replace 256 with MEDIA_FLAGS.MEDIA_AUDIO in version 1.3
            End If
        Next
        nFadeVolume = 0
        If Flags And MEDIA_FLAGS.MEDIA_SUSPEND_OTHERS Then
            strFadeToMedia = strPlayingMedia
            strPlayingMedia = ""
            Flags = Flags Or MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND
            Exit Sub
        End If
    End If
    
    If FGM Is Nothing Then
        strPlayingMedia = ""
        Exit Sub
    End If
    
    Set MPos = FGM
    
    If MPos.CurrentPosition >= MPos.Duration Then
        If (Flags And MEDIA_FLAGS.MEDIA_LOOP) = 0 Then
            StopClip True
            Exit Sub
        End If
        MPos.CurrentPosition = 0
    End If
    
    If Flags And MEDIA_FLAGS.MEDIA_VIDEO Then
        Set VidWin = FGM
        Set VidInf = FGM
        If VidWin.Visible = False Then
            VidWin.WindowStyle = WS_VISIBLE Or WS_CHILD
            VidWin.SetWindowPosition OutputX, OutputY, VidInf.VideoWidth, VidInf.VideoHeight
            If Not (CurDisp Is Nothing) Then
                CurDisp.FlipToGDISurface
                VidWin.Owner = CurDisp.hwnd
            End If
            VidWin.Visible = True
        End If
    End If

    FGM.GetState 1000, nState
    If nState = 0 Then FGM.Run
    If Flags And 256 Then ' Replace 256 with MEDIA_FLAGS.MEDIA_AUDIO in version 1.3
        Set AudInf = FGM
        If AudInf.Volume <> nFadeVolume Then AudInf.Volume = nFadeVolume
    End If

    If Flags And MEDIA_FLAGS.MEDIA_MODAL Then
        bCancelModalMedia = False
        Do
            Sleep 100
            If Not (Prj Is Nothing) Then
                If Not (Prj.GamePlayer Is Nothing) Then
                    If Prj.GamePlayer.bEnableJoystick Then
                        Prj.GamePlayer.PollJoystick JOYSTICKID1
                        If Prj.GamePlayer.JSValid Then Prj.GamePlayer.JoyMove
                    End If
                End If
            End If
            DoEvents
        Loop While MPos.CurrentPosition < MPos.Duration And bCancelModalMedia = False
        If (Flags And MEDIA_FLAGS.MEDIA_KEEP_LOADED) = 0 Then
            Set FGM = Nothing
        End If
        strPlayingMedia = ""
        Exit Sub
    End If

    Exit Sub
    
PlayErr:
    strErr = Err.Description
    If Not CurDisp Is Nothing Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    MsgBox strErr, vbExclamation
End Sub

Public Sub StopClip(ByVal bForceRewind As Boolean)
Attribute StopClip.VB_Description = "Abruptly stop playing the clip."
    Dim MPos As IMediaPosition
    
    On Error GoTo StopErr
    
    strPlayingMedia = ""
    If FGM Is Nothing Then Exit Sub
    FGM.Stop
    If (Flags And MEDIA_FLAGS.MEDIA_KEEP_LOADED) = 0 Then
        Set FGM = Nothing
        strLoadedMedia = ""
    End If
    
    If bForceRewind And Not (FGM Is Nothing) Then
        Set MPos = FGM
        MPos.CurrentPosition = 0
    End If
    
    Exit Sub
StopErr:
    If Not CurDisp Is Nothing Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Reset()
Attribute Reset.VB_Description = "Reset this clip to the state it was in when the project was loaded."
    Set FGM = Nothing
    strPlayingMedia = ""
    strLoadedMedia = ""
    strFadeToMedia = ""
End Sub

Public Function IsPlaying() As Boolean
Attribute IsPlaying.VB_Description = "Returns True if the clip is currently playing."
    IsPlaying = (Len(strPlayingMedia) > 0)
End Function

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the definition of this clip to a string for the project file."
    Serialize = GeneralSerialize("Name", "MediaFile", "OutputX", "OutputY", "Flags", "Volume", _
        Name, strMediaFile, OutputX, OutputY, Flags, Volume)
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a media clip definition from the project file (string)."
    GeneralDeserialize Data, "Name", "MediaFile", "OutputX", "OutputY", "Flags", "Volume", _
        Name, strMediaFile, OutputX, OutputY, Flags, Volume
End Sub
