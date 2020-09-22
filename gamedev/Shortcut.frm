VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShortcut 
   Caption         =   "Make Shortcut"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   HelpContextID   =   124
   Icon            =   "Shortcut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowseScript 
      Caption         =   "..."
      Height          =   285
      Left            =   2205
      TabIndex        =   13
      Top             =   2400
      Width           =   330
   End
   Begin VB.TextBox txtScriptPath 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame fraScreenDepth 
      Caption         =   "Display Mode"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   2415
      Begin VB.OptionButton optColorDepth 
         Caption         =   "32 bits per pixel"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optColorDepth 
         Caption         =   "24 bits per pixel"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optColorDepth 
         Caption         =   "16 bits per pixel"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optColorDepth 
         Caption         =   "Use GameDev default"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.TextBox txtGamePath 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtGameDevPath 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtShortcutName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cdlScriptPath 
      Left            =   2280
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".vbs"
      DialogTitle     =   "Browse for Script File"
      Filter          =   "All Files (*.*)|*.*|VBScript Files (*.vbs)|*.vbs"
      FilterIndex     =   2
      Flags           =   4100
   End
   Begin VB.Label lblScriptPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Path to script (blank for none):"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblGamePath 
      BackStyle       =   0  'Transparent
      Caption         =   "Path to game project:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblGameDevPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Path to GameDev.exe"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblLnkName 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmShortcut"
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
' File: Shortcut.frm - Make Shortcut Dialog
'
'======================================================================

Option Explicit

Private Sub cmdBrowseScript_Click()
    On Error Resume Next
    If Len(Prj.ProjectPath) Then cdlScriptPath.InitDir = PathFromFile(Prj.ProjectPath)
    Err.Clear
    cdlScriptPath.ShowOpen
    If Err.Number = 0 Then
        txtScriptPath.Text = cdlScriptPath.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lReturn As Long
    Dim strArgs As String
    Dim strErr As String
    Dim strLnkPath As String
    Dim strShortcutName As String
    Dim strGameDevPath As String
    Dim strParent As String
    
    On Error Resume Next
    If InStr(txtGamePath.Text, " ") > 0 Then
        If Left$(txtGamePath.Text, 1) <> """" Then
            txtGamePath.Text = """" & txtGamePath.Text & """"
        End If
    End If
    
    If InStr(txtScriptPath.Text, " ") > 0 Then
        If Left$(txtScriptPath.Text, 1) <> """" Then
            txtScriptPath.Text = """" & txtScriptPath.Text & """"
        End If
    End If
    
    strArgs = txtGamePath & " /p"
    If Len(Trim$(txtScriptPath.Text)) Then strArgs = strArgs & " " & txtScriptPath.Text
    If optColorDepth(1).Value Then strArgs = strArgs & " /d 16"
    If optColorDepth(2).Value Then strArgs = strArgs & " /d 24"
    If optColorDepth(3).Value Then strArgs = strArgs & " /d 32"

    strArgs = strArgs & Chr$(0)
    strLnkPath = "..\..\Desktop" & Chr$(0)
    strShortcutName = txtShortcutName.Text & Chr$(0)
    strGameDevPath = txtGameDevPath.Text & Chr$(0)
    strParent = "$(Programs)" & Chr$(0)

    If Err.Number <> 0 Then
        If MsgBox("There was an error gathering parameters from the dialog: " & vbCrLf & Err.Description & vbCrLf & "Continue?", vbExclamation + vbYesNo + vbDefaultButton2, "Creating Shortcut") <> vbYes Then Exit Sub
        Err.Clear
    End If

    lReturn = VB6fCreateShellLink(strLnkPath, strShortcutName, strGameDevPath, strArgs, True, strParent)
    If (Err.Number <> 0) Or (lReturn = 0) Then
        If Err.Number <> 0 Then strErr = Err.Description Else strErr = "Error number " & Err.LastDllError
        strErr = vbCrLf & "(" & strErr & ")"
        If MsgBox("Failed to create shortcut using VB6STKIT.DLL, try again with VB5STKIT.DLL?" & strErr, vbExclamation + vbYesNo) = vbYes Then
            Err.Clear
            lReturn = VB5fCreateShellLink(strLnkPath, strShortcutName, strGameDevPath, strArgs)
            If (Err.Number <> 0) Or lReturn = 0 Then
                MsgBox "Failed to create shortcut. " & Err.Description, vbExclamation
            Else
                MsgBox "Succeeded.  The shortcut was created on the desktop.", vbInformation
                Unload Me
            End If
        End If
    Else
        MsgBox "Shortcut """ & txtShortcutName.Text & """ has been created on the desktop", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "MakeShortcut", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    txtGameDevPath.Text = App.Path & "\" & App.EXEName & ".exe"
    txtGamePath.Text = Prj.ProjectPath
    
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "MakeShortcut", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

