VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multimedia Manager"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   HelpContextID   =   110
   Icon            =   "Media.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMediaClips 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      Begin MSComCtlLib.Slider sldVolume 
         Height          =   495
         Left            =   2880
         TabIndex        =   38
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   500
         SmallChange     =   100
         Min             =   -10000
         Max             =   0
         TickFrequency   =   500
      End
      Begin VB.CommandButton cmdUpdateClip 
         Caption         =   "&Update Clip"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteClip 
         Caption         =   "&Delete Clip"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdAddClip 
         Caption         =   "&Add Clip"
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame fraVideo 
         Caption         =   "Video Settings"
         Height          =   615
         Left            =   2160
         TabIndex        =   13
         Top             =   2520
         Width           =   2415
         Begin VB.TextBox txtMediaY 
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtMediaX 
            Height          =   285
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblMediaX 
            Caption         =   "X:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblMediaY 
            Caption         =   "Y:"
            Height          =   255
            Left            =   1320
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.CheckBox chkSuspendOther 
         Caption         =   "Suspend other media to play"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkModal 
         Caption         =   "Modal"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkLoop 
         Caption         =   "Play looping"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chkKeepLoaded 
         Caption         =   "Keep loaded (short/fast clip)"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmdMediaBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   4350
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtClipName 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtMediaFile 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.ListBox lstMediaClips 
         Height          =   2790
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblVolume 
         Caption         =   "Volume:"
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblClipName 
         BackStyle       =   0  'Transparent
         Caption         =   "Clip name:"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblMediaFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Media file:"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblMediaClips 
         BackStyle       =   0  'Transparent
         Caption         =   "Media clips:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame fraAssignMedia 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   4695
      Begin MSComDlg.CommonDialog cdlBrowse 
         Left            =   120
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "Common ActiveMovie Media (*.wav;*.mp3;*.avi;*.mpg;*.mid)|*.wav;*.mp3;*.avi;*.mpg;*.mid|All Files (*.*)|*.*"
         FilterIndex     =   1
         Flags           =   4108
      End
      Begin VB.ComboBox cboBGMusic 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboMaps 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Width           =   2055
      End
      Begin VB.Frame fraFuncMedia 
         Caption         =   "Special Functions"
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   4455
         Begin VB.OptionButton optFuncStop 
            Caption         =   "Stop clip"
            Height          =   255
            Left            =   3360
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdBrowseBG 
            Caption         =   "..."
            Height          =   270
            Left            =   4110
            TabIndex        =   36
            Top             =   1800
            Width           =   255
         End
         Begin VB.TextBox txtBGMusic 
            Height          =   285
            Left            =   2550
            TabIndex        =   35
            Top             =   1800
            Width           =   1545
         End
         Begin VB.OptionButton optChangeBGMusic 
            Caption         =   "Change music"
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton optFuncClip 
            Caption         =   "Play clip"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optFuncNone 
            Caption         =   "No media"
            Height          =   255
            Left            =   2280
            TabIndex        =   28
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cboFuncMedia 
            Height          =   315
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   960
            Width           =   1785
         End
         Begin VB.ListBox lstFunctions 
            Height          =   2010
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblNewMediaFile 
            Caption         =   "New media file:"
            Height          =   255
            Left            =   2550
            TabIndex        =   34
            Top             =   1560
            Width           =   1785
         End
         Begin VB.Label lblFuncMedia 
            Caption         =   "Function media clip:"
            Height          =   255
            Left            =   2550
            TabIndex        =   31
            Top             =   720
            Width           =   1785
         End
      End
      Begin VB.Label lblBGMusic 
         Caption         =   "Background music:"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblMaps 
         Caption         =   "Maps:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1935
      End
   End
   Begin MSComCtlLib.TabStrip TabMedia 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7435
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Media Clips"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Assign Media"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManageMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Media.frm - Multimedia Manager Dialog
'
'======================================================================

Option Explicit

Dim bLoadFuncMedia As Boolean

Private Sub cboBGMusic_Change()
    If cboMaps.ListIndex >= 0 Then UpdateFuncMedia
End Sub

Private Sub cboBGMusic_Click()
    cboBGMusic_Change
End Sub

Private Sub cboFuncMedia_Change()
    If bLoadFuncMedia Then Exit Sub
    UpdateFuncMedia
End Sub

Private Sub cboFuncMedia_Click()
    If bLoadFuncMedia Then Exit Sub
    UpdateFuncMedia
End Sub

Private Sub cboMaps_Change()
    Dim Idx As Integer
    
    On Error GoTo LoadFuncErr
    
    lstFunctions.Clear
    With Prj.Maps(cboMaps.List(cboMaps.ListIndex))
        For Idx = 0 To .SpecialCount - 1
            lstFunctions.AddItem .Specials(Idx).Name
        Next
        For Idx = 0 To cboBGMusic.ListCount - 1
            If cboBGMusic.List(Idx) = .BackgroundMusic Then
                cboBGMusic.ListIndex = Idx
                Exit For
            End If
        Next
    End With

    Exit Sub

LoadFuncErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cboMaps_Click()
    cboMaps_Change
End Sub

Private Sub cmdAddClip_Click()
    Dim Clp As New MediaClip
    
    If Prj.MediaMgr.ClipExists(txtClipName.Text) Then
        MsgBox "Please choose a unique name for a new media clip."
        Exit Sub
    End If
    
    UpdateClip Clp
    Prj.MediaMgr.AddClip Clp
    LoadMediaLists
End Sub

Private Sub cmdBrowseBG_Click()
    cdlBrowse.DialogTitle = "Specify New Background Music"
    If Len(Prj.ProjectPath) Then
        cdlBrowse.FileName = GetRelativePath(Prj.ProjectPath, txtBGMusic.Text)
        cdlBrowse.InitDir = PathFromFile(txtBGMusic.Text)
    End If
    On Error Resume Next
    cdlBrowse.ShowOpen
    If Err.Number Then Exit Sub
    txtBGMusic.Text = GetRelativePath(Prj.ProjectPath, cdlBrowse.FileName)
End Sub

Private Sub cmdDeleteClip_Click()
    If lstMediaClips.ListIndex < 0 Then
        MsgBox "Please select a clip to delete first.", vbExclamation
        Exit Sub
    End If
    
    Prj.MediaMgr.RemoveClip lstMediaClips.List(lstMediaClips.ListIndex)
    LoadMediaLists
End Sub

Private Sub cmdMediaBrowse_Click()
    cdlBrowse.DialogTitle = "Locate Media Clip File"
    On Error Resume Next
    If Len(Prj.ProjectPath) Then
        cdlBrowse.FileName = GetRelativePath(Prj.ProjectPath, txtMediaFile.Text)
        cdlBrowse.InitDir = PathFromFile(Prj.ProjectPath)
    End If
    cdlBrowse.ShowOpen
    If Err.Number Then Exit Sub
    txtMediaFile.Text = GetRelativePath(Prj.ProjectPath, cdlBrowse.FileName)
End Sub

Private Sub cmdUpdateClip_Click()
    On Error Resume Next
    If lstMediaClips.ListIndex < 0 Then
        MsgBox "Please select a clip to update first.", vbExclamation
        Exit Sub
    End If
    UpdateClip Prj.MediaMgr.Clip(lstMediaClips.List(lstMediaClips.ListIndex))
    If Err.Number Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    Dim Idx As Integer
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "MediaManager", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6)), CLng(Mid$(WndPos, 15, 6)), CLng(Right$(WndPos, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
        
    cboMaps.Clear
    For Idx = 0 To Prj.MapCount - 1
        cboMaps.AddItem Prj.Maps(Idx).Name
    Next
    
    LoadMediaLists
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
    
End Sub

Sub LoadMediaLists()
    Dim Idx As Integer
    
    cboBGMusic.Clear
    cboFuncMedia.Clear
    lstMediaClips.Clear
    For Idx = 0 To Prj.MediaMgr.MediaClipCount - 1
        cboBGMusic.AddItem Prj.MediaMgr.Clip(Idx).Name
        cboFuncMedia.AddItem Prj.MediaMgr.Clip(Idx).Name
        lstMediaClips.AddItem Prj.MediaMgr.Clip(Idx).Name
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UpdateFuncMedia
    SaveSetting "GameDev", "Windows", "MediaManager", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000") & "," & Format$(Me.Width, " 00000;-00000") & "," & Format$(Me.Height, " 00000;-00000")
End Sub

Private Sub lstFunctions_Click()
    LoadFuncMedia
End Sub

Sub LoadFuncMedia()
    On Error GoTo LoadFuncMediaErr
    If cboMaps.ListIndex < 0 Or lstFunctions.ListIndex < 0 Then
        optFuncNone.Value = True
        cboFuncMedia.ListIndex = -1
        cboFuncMedia.Enabled = False
        txtBGMusic.Text = ""
        txtBGMusic.Enabled = False
        cmdBrowseBG.Enabled = False
        Exit Sub
    End If
    bLoadFuncMedia = True
    With Prj.Maps(cboMaps.List(cboMaps.ListIndex)).Specials(lstFunctions.List(lstFunctions.ListIndex))
        If Len(.MediaName) = 0 Then
            optFuncNone.Value = True
            cboFuncMedia.ListIndex = -1
            cboFuncMedia.Enabled = False
            txtBGMusic.Text = ""
            txtBGMusic.Enabled = False
            cmdBrowseBG.Enabled = False
        ElseIf .Flags And InteractionFlags.INTFL_STOPMEDIA Then
            SelectFuncMedia .MediaName
            optFuncStop.Value = True
        ElseIf .Flags And InteractionFlags.INTFL_CHANGEBGMEDIA Then
            txtBGMusic.Text = .MediaName
            cboFuncMedia.Enabled = False
            cboFuncMedia.ListIndex = -1
            txtBGMusic.Enabled = True
            cmdBrowseBG.Enabled = True
            optChangeBGMusic.Value = True
        Else
            SelectFuncMedia .MediaName
            optFuncClip.Value = True
        End If
    End With
    bLoadFuncMedia = False
    Exit Sub
    
LoadFuncMediaErr:
    bLoadFuncMedia = False
    MsgBox Err.Description, vbExclamation
End Sub

Sub UpdateFuncMedia()
    On Error GoTo UpdateFuncMediaErr
    If cboMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(cboMaps.List(cboMaps.ListIndex))
        Prj.IsDirty = True
        .BackgroundMusic = cboBGMusic.Text
        If lstFunctions.ListIndex >= 0 Then
            With .Specials(lstFunctions.List(lstFunctions.ListIndex))
                If optFuncNone.Value Then
                    .MediaName = ""
                    .Flags = .Flags And Not (InteractionFlags.INTFL_CHANGEBGMEDIA Or InteractionFlags.INTFL_STOPMEDIA)
                ElseIf optFuncStop.Value Then
                    .MediaName = cboFuncMedia.Text
                    .Flags = (.Flags And Not InteractionFlags.INTFL_CHANGEBGMEDIA) Or InteractionFlags.INTFL_STOPMEDIA
                ElseIf optChangeBGMusic.Value Then
                    .MediaName = txtBGMusic.Text
                    .Flags = (.Flags And Not InteractionFlags.INTFL_STOPMEDIA) Or InteractionFlags.INTFL_CHANGEBGMEDIA
                ElseIf optFuncClip.Value Then
                    .MediaName = cboFuncMedia.Text
                    .Flags = .Flags And Not (InteractionFlags.INTFL_CHANGEBGMEDIA Or InteractionFlags.INTFL_STOPMEDIA)
                End If
            End With
        End If
    End With

    Exit Sub
    
UpdateFuncMediaErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub SelectFuncMedia(MediaName As String)
    Dim Idx As Integer
    
    cboFuncMedia.Enabled = True
    txtBGMusic.Text = ""
    txtBGMusic.Enabled = False
    cmdBrowseBG.Enabled = False
    For Idx = 0 To cboFuncMedia.ListCount - 1
        If cboFuncMedia.List(Idx) = MediaName Then
            cboFuncMedia.ListIndex = Idx
            Exit Sub
        End If
    Next
    cboFuncMedia.ListIndex = -1
End Sub

Private Sub lstMediaClips_Click()
    On Error Resume Next
    LoadClip Prj.MediaMgr.Clip(lstMediaClips.List(lstMediaClips.ListIndex))
    If Err.Number Then MsgBox Err.Description, vbExclamation
End Sub

Sub LoadClip(Clp As MediaClip)
    On Error GoTo LoadClipErr
    With Clp
        txtClipName.Text = .Name
        txtMediaFile.Text = .strMediaFile
        chkKeepLoaded.Value = IIf(.Flags And MEDIA_FLAGS.MEDIA_KEEP_LOADED, vbChecked, vbUnchecked)
        chkLoop.Value = IIf(.Flags And MEDIA_FLAGS.MEDIA_LOOP, vbChecked, vbUnchecked)
        chkModal.Value = IIf(.Flags And MEDIA_FLAGS.MEDIA_MODAL, vbChecked, vbUnchecked)
        chkSuspendOther.Value = IIf(.Flags And MEDIA_FLAGS.MEDIA_SUSPEND_OTHERS, vbChecked, vbUnchecked)
        txtMediaX.Text = CStr(.OutputX)
        txtMediaY.Text = CStr(.OutputY)
        sldVolume.Value = .Volume
    End With
    Exit Sub
LoadClipErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub optChangeBGMusic_Click()
    If bLoadFuncMedia Then Exit Sub
    UpdateFuncMedia
    If cboMaps.ListIndex >= 0 Then
        With Prj.Maps(cboMaps.List(cboMaps.ListIndex))
            If Prj.MediaMgr.ClipExists(.BackgroundMusic) Then
                txtBGMusic.Text = Prj.MediaMgr.Clip(.BackgroundMusic).strMediaFile
            Else
                MsgBox "Cannot change background music unless map has background music.", vbExclamation
            End If
        End With
    End If
    LoadFuncMedia
End Sub

Private Sub optFuncClip_Click()
    If bLoadFuncMedia Then Exit Sub
    If cboFuncMedia.ListCount > 0 Then If cboFuncMedia.ListIndex < 0 Then cboFuncMedia.ListIndex = 0
    UpdateFuncMedia
    LoadFuncMedia
End Sub

Private Sub optFuncNone_Click()
    If bLoadFuncMedia Then Exit Sub
    UpdateFuncMedia
    LoadFuncMedia
End Sub

Private Sub optFuncStop_Click()
    If bLoadFuncMedia Then Exit Sub
    If cboFuncMedia.ListCount > 0 Then If cboFuncMedia.ListIndex < 0 Then cboFuncMedia.ListIndex = 0
    UpdateFuncMedia
    LoadFuncMedia
End Sub

Private Sub TabMedia_Click()
    fraMediaClips.Visible = (TabMedia.SelectedItem.Index = 1)
    fraAssignMedia.Visible = (TabMedia.SelectedItem.Index = 2)
    UpdateFuncMedia
End Sub

Sub UpdateClip(Clp As MediaClip)
    On Error GoTo UpdateClipErr
    
    With Clp
        Prj.IsDirty = True
        .Name = txtClipName.Text
        .strMediaFile = txtMediaFile.Text
        .Flags = IIf(chkKeepLoaded.Value = vbChecked, MEDIA_FLAGS.MEDIA_KEEP_LOADED, 0)
        If chkLoop.Value = vbChecked Then .Flags = .Flags Or MEDIA_FLAGS.MEDIA_LOOP
        If chkModal.Value = vbChecked Then .Flags = .Flags Or MEDIA_FLAGS.MEDIA_MODAL
        If chkSuspendOther.Value = vbChecked Then .Flags = .Flags Or MEDIA_FLAGS.MEDIA_SUSPEND_OTHERS
        .OutputX = Val(txtMediaX.Text)
        .OutputY = Val(txtMediaY.Text)
        .Volume = sldVolume.Value
    End With
    Exit Sub
    
UpdateClipErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub txtBGMusic_Change()
    If bLoadFuncMedia Then Exit Sub
    UpdateFuncMedia
End Sub
