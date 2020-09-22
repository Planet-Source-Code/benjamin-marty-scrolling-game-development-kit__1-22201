VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "GameDev Options"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   HelpContextID   =   123
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBackupForPlay 
      Caption         =   "Store project backup in memory while playing"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame fraDisplayMode 
      Caption         =   "Default Full Screen Display Mode"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   3855
      Begin VB.OptionButton optColorDepth 
         Caption         =   "32 bits per pixel"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optColorDepth 
         Caption         =   "24 bits per pixel"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optColorDepth 
         Caption         =   "16 bits per pixel"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CheckBox chkPlayWarn 
      Caption         =   "Display warning message before playing map."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CheckBox chkDisablePlayerEdit 
      Caption         =   "Disable player edits in map editor."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "(Turning this on prevents the player sprite from causing permanent changes while editing the map.)"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Options.frm - GameDev Options Dialog
'
'======================================================================

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ScreenDepth As Integer

    On Error GoTo OptOKErr

    SaveSetting "GameDev", "Options", "DisablePlayerEdit", IIf(chkDisablePlayerEdit.Value = vbChecked, "1", "0")
    SaveSetting "GameDev", "Options", "PlayWarn", IIf(chkPlayWarn.Value = vbChecked, "1", "0")
    SaveSetting "GameDev", "Options", "BackupForPlay", IIf(chkBackupForPlay.Value = vbChecked, "1", "0")
    ScreenDepth = 16
    If optColorDepth(1).Value Then ScreenDepth = 24
    If optColorDepth(2).Value Then ScreenDepth = 32
    SaveSetting "GameDev", "Options", "ScreenDepth", CStr(ScreenDepth)
    
    Unload Me
    
    Exit Sub

OptOKErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    
    On Error GoTo OptLoadErr
    
    chkDisablePlayerEdit.Value = IIf(Val(GetSetting("GameDev", "Options", "DisablePlayerEdit", "0")) <> 0, vbChecked, vbUnchecked)
    chkPlayWarn.Value = IIf(GetSetting("GameDev", "Options", "PlayWarn", "1") <> 0, vbChecked, vbUnchecked)
    chkBackupForPlay.Value = IIf(GetSetting("GameDev", "Options", "BackupForPlay", "1") <> 0, vbChecked, vbUnchecked)
    Select Case Val(GetSetting("GameDev", "Options", "ScreenDepth", "16"))
    Case 24
        optColorDepth(1).Value = True
    Case 32
        optColorDepth(2).Value = True
    Case Else
        optColorDepth(0).Value = True
    End Select
    
    Exit Sub
    
OptLoadErr:
    MsgBox Err.Description, vbExclamation

End Sub
