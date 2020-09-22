VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Manages multimedia clips for the project.  Owned by and saved with the project."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MediaMgr.cls - Multimedia Manager Class
'
'======================================================================

Option Explicit

Dim MediaClips() As MediaClip
Dim ClipCount As Integer

Public Function NewClip(ByVal bAddToCollection As Boolean, Optional ByVal ClipName, Optional ByVal MediaFile, _
                        Optional ByVal Flags, Optional ByVal OutputX, Optional ByVal OutputY, Optional ByVal Volume As Integer) As MediaClip
Attribute NewClip.VB_Description = "Create and return a new MediaClip with the specified parameters, optionally adding it to the manager."
    Dim Clp As New MediaClip
    
    If Not IsMissing(ClipName) Then Clp.Name = ClipName
    If Not IsMissing(MediaFile) Then Clp.strMediaFile = MediaFile
    If Not IsMissing(Flags) Then Clp.Flags = Flags
    If Not IsMissing(OutputX) Then Clp.OutputX = OutputX
    If Not IsMissing(OutputY) Then Clp.OutputY = OutputY
    If Not IsMissing(Volume) Then Clp.Volume = Volume
    If bAddToCollection Then AddClip Clp
    Set NewClip = Clp
End Function

Public Sub AddClip(Clip As MediaClip)
Attribute AddClip.VB_Description = "Add a clip object to the manager to be managed / saved with the project."
    If ClipCount > 0 Then
        ReDim Preserve MediaClips(0 To ClipCount) As MediaClip
        ClipCount = ClipCount + 1
    Else
        ReDim MediaClips(0) As MediaClip
        ClipCount = 1
    End If
    
    Set MediaClips(UBound(MediaClips)) = Clip
End Sub

Public Sub RemoveClip(ByVal Key As Variant)
Attribute RemoveClip.VB_Description = "Remove a MediaClip from the MediaManager by name or index."
    Dim Idx As Integer
    
    If IsNumeric(Key) Then
        Idx = CInt(Key)
    Else
        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).Name = Key Then Exit For
        Next
    End If
    
    If Idx < 0 Or Idx >= ClipCount Then
        Err.Raise vbObjectError, , "RemoveClip cannot find element " & Key
    End If
    
    ClipCount = ClipCount - 1
    Do While Idx < ClipCount
        Set MediaClips(Idx) = MediaClips(Idx + 1)
        Idx = Idx + 1
    Loop
    
    Set MediaClips(Idx) = Nothing
    
    If ClipCount > 0 Then
        ReDim Preserve MediaClips(0 To ClipCount - 1) As MediaClip
    Else
        Erase MediaClips
    End If
    
End Sub

Public Property Get Clip(ByVal Key As Variant) As MediaClip
Attribute Clip.VB_Description = "Return a MediaClip object given its name or ordinal index in the MediaManager."
Attribute Clip.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Clip.VB_UserMemId = 0
    Dim Idx As Integer
    
    If IsNumeric(Key) Then
        Idx = CInt(Key)
    Else
        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).Name = Key Then Exit For
        Next
    End If
    
    If Idx < 0 Or Idx >= ClipCount Then
        Err.Raise vbObjectError, , "Cannot find media clip " & Key
    End If
    
    Set Clip = MediaClips(Idx)
End Property

Public Function ClipExists(Key As Variant) As Boolean
Attribute ClipExists.VB_Description = "Return True if a MediaClip by the specified name or index exists in the MediaManager."
    Dim Idx As Integer

    If IsNumeric(Key) Then
        Idx = CInt(Key)
    Else
        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).Name = Key Then Exit For
        Next
    End If
    
    If Idx >= 0 And Idx < ClipCount Then
        ClipExists = True
    Else
        ClipExists = False
    End If
    
End Function

Public Property Get MediaClipCount() As Integer
Attribute MediaClipCount.VB_Description = "Return the number of MediaClips in the MediaManager."
    MediaClipCount = ClipCount
End Property

Public Sub ModalFadeOutAll()
Attribute ModalFadeOutAll.VB_Description = "Fade out and stop all media clips, then return to the caller."
    Dim Idx As Integer
    Dim nPlayingCount As Integer
    
    For Idx = 0 To ClipCount - 1
        If MediaClips(Idx).IsPlaying Then MediaClips(Idx).FadeOut
    Next
    
    Do
        RenderActiveMedia
        nPlayingCount = 0
        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).IsPlaying Then nPlayingCount = nPlayingCount + 1
        Next
        Sleep 50
    Loop Until nPlayingCount = 0
End Sub

Public Sub CancelModalMedia()
Attribute CancelModalMedia.VB_Description = "Stop playing modal clips now (in response to an event)."
    Dim Idx As Integer
    
    For Idx = 0 To ClipCount - 1
        If MediaClips(Idx).IsPlaying Then
            If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_SUSPEND_OTHERS Then
                MediaClips(Idx).bCancelModalMedia = True
            End If
        End If
    Next
End Sub

Public Sub RenderActiveMedia()
Attribute RenderActiveMedia.VB_Description = "Continue processing all MediaClips according to their various processes and parameters."
    Dim Idx As Integer
    Dim nSuspendCount As Integer
    Dim nSuspendedCount As Integer
    Dim nPlayingCount As Integer
    Dim bSuspendedPlaying As Boolean
    Dim bModalLoop As Boolean
   
    Do
        nSuspendCount = 0
        nSuspendedCount = 0
        nPlayingCount = 0
        bSuspendedPlaying = False
        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND Then
                nSuspendCount = nSuspendCount + 1
                If MediaClips(Idx).IsPlaying Then
                    nPlayingCount = nPlayingCount + 1
                    bSuspendedPlaying = True
                End If
            Else
                If MediaClips(Idx).IsPlaying Then nPlayingCount = nPlayingCount + 1
            End If
            If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_SUSPEND Then nSuspendedCount = nSuspendedCount + 1
        Next
    
        If nPlayingCount = 0 And nSuspendCount = 0 Then
            If nSuspendedCount > 0 Then
                For Idx = 0 To ClipCount - 1
                    If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_SUSPEND Then
                        MediaClips(Idx).Flags = MediaClips(Idx).Flags And Not MEDIA_FLAGS.MEDIA_SUSPEND
                        MediaClips(Idx).Play
                    End If
                Next
            End If
            bModalLoop = False
        End If

        For Idx = 0 To ClipCount - 1
            If MediaClips(Idx).IsPlaying Then MediaClips(Idx).Play
            If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_WAITING_ON_SUSPEND Then
                If nPlayingCount = 0 And Not bSuspendedPlaying Then MediaClips(Idx).Play
                If MediaClips(Idx).Flags And MEDIA_FLAGS.MEDIA_MODAL Then bModalLoop = True
            ElseIf nSuspendCount > 0 And MediaClips(Idx).IsPlaying Then
                MediaClips(Idx).Flags = MediaClips(Idx).Flags Or MEDIA_FLAGS.MEDIA_SUSPEND
                MediaClips(Idx).FadeOut
            End If
        Next
        If bModalLoop Then Sleep 50
    Loop While bModalLoop

End Sub

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store all MediaClip definitions into a string for the project file."
    Dim Rtn As String
    Dim Idx As Integer
    
    Rtn = GeneralSerialize("MediaClipCount", ClipCount)
    For Idx = 0 To ClipCount - 1
        Rtn = Rtn & Clip(Idx).Serialize
    Next
    Serialize = Rtn
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load all the clips for the project into the MediaManager from a string."
    Dim Idx As Integer
    
    GeneralDeserialize Data, "MediaClipCount", ClipCount
    If ClipCount = 0 Then
        Erase MediaClips
        Exit Sub
    Else
        ReDim MediaClips(0 To ClipCount - 1)
    End If
    For Idx = 0 To ClipCount - 1
        Set MediaClips(Idx) = New MediaClip
        MediaClips(Idx).Deserialize Data
    Next
End Sub

Public Sub UnloadAll()
Attribute UnloadAll.VB_Description = "Reset all MediaClips to their initial state."
    Dim Idx As Integer
    
    For Idx = 0 To ClipCount - 1
        MediaClips(Idx).Reset
    Next
End Sub
