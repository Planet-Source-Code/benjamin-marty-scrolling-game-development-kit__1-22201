VERSION 5.00
Begin VB.Form frmCtrlConfig 
   Caption         =   "Controller Configuration"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   HelpContextID   =   102
   Icon            =   "CtrlCnfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnableJoystick 
      Caption         =   "Enable Joystick"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtBtn4 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtBtn3 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtBtn2 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtBtn1 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtDown 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtRight 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtUp 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Down"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label lblRight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Up"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBtn4 
      BackStyle       =   0  'Transparent
      Caption         =   "Button 4:"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblBtn3 
      BackStyle       =   0  'Transparent
      Caption         =   "Button 3:"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblBtn2 
      BackStyle       =   0  'Transparent
      Caption         =   "Button 2:"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblBtn1 
      BackStyle       =   0  'Transparent
      Caption         =   "Button 1:"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCtrlConfig"
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
' File: CtrlCnfg.frm - Controller Configuration Dialog
'
'======================================================================

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Prj.GamePlayer.KeyConfig(0) = CInt(txtUp.Tag)
    Prj.GamePlayer.KeyConfig(1) = CInt(txtLeft.Tag)
    Prj.GamePlayer.KeyConfig(2) = CInt(txtRight.Tag)
    Prj.GamePlayer.KeyConfig(3) = CInt(txtDown.Tag)
    Prj.GamePlayer.KeyConfig(4) = CInt(txtBtn1.Tag)
    Prj.GamePlayer.KeyConfig(5) = CInt(txtBtn2.Tag)
    Prj.GamePlayer.KeyConfig(6) = CInt(txtBtn3.Tag)
    Prj.GamePlayer.KeyConfig(7) = CInt(txtBtn4.Tag)
    Prj.GamePlayer.bEnableJoystick = (chkEnableJoystick.Value = vbChecked)
    Prj.IsDirty = True
    Unload Me
End Sub

Private Sub Form_Load()
    HandleKeyDown Prj.GamePlayer.KeyConfig(0), txtUp
    HandleKeyDown Prj.GamePlayer.KeyConfig(1), txtLeft
    HandleKeyDown Prj.GamePlayer.KeyConfig(2), txtRight
    HandleKeyDown Prj.GamePlayer.KeyConfig(3), txtDown
    HandleKeyDown Prj.GamePlayer.KeyConfig(4), txtBtn1
    HandleKeyDown Prj.GamePlayer.KeyConfig(5), txtBtn2
    HandleKeyDown Prj.GamePlayer.KeyConfig(6), txtBtn3
    HandleKeyDown Prj.GamePlayer.KeyConfig(7), txtBtn4
    chkEnableJoystick.Value = IIf(Prj.GamePlayer.bEnableJoystick, vbChecked, vbUnchecked)
End Sub

Private Sub txtBtn1_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtBtn1
End Sub

Private Sub txtBtn2_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtBtn2
End Sub

Private Sub txtBtn3_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtBtn3
End Sub

Private Sub txtBtn4_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtBtn4
End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtDown
End Sub

Private Sub txtLeft_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtLeft
End Sub

Private Sub txtRight_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtRight
End Sub

Private Sub txtUp_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, txtUp
End Sub

Sub HandleKeyDown(KeyCode As Integer, KeyControl As TextBox)
    Dim KeyCodeName As String
    
    Select Case KeyCode
    Case vbKeyAdd
        KeyCodeName = "(Num)+"
    Case vbKeyBack
        KeyCodeName = "Backspace"
    Case vbKeyControl
        KeyCodeName = "Control"
    Case vbKeyDecimal
        KeyCodeName = "(Num)."
    Case vbKeyDelete
        KeyCodeName = "Delete"
    Case vbKeyDivide
        KeyCodeName = "(Num)/"
    Case vbKeyDown
        KeyCodeName = "Down"
    Case vbKeyEnd
        KeyCodeName = "End"
    Case vbKeyEscape
        KeyCodeName = "Escape"
    Case vbKeyF1
        KeyCodeName = "F1"
    Case vbKeyF2
        KeyCodeName = "F2"
    Case vbKeyF3
        KeyCodeName = "F3"
    Case vbKeyF4
        KeyCodeName = "F4"
    Case vbKeyF5
        KeyCodeName = "F5"
    Case vbKeyF6
        KeyCodeName = "F6"
    Case vbKeyF7
        KeyCodeName = "F7"
    Case vbKeyF8
        KeyCodeName = "F8"
    Case vbKeyF9
        KeyCodeName = "F9"
    Case vbKeyF10
        KeyCodeName = "F10"
    Case vbKeyF11
        KeyCodeName = "F11"
    Case vbKeyF12
        KeyCodeName = "F12"
    Case vbKeyHome
        KeyCodeName = "Home"
    Case vbKeyInsert
        KeyCodeName = "Insert"
    Case vbKeyLeft
        KeyCodeName = "Left"
    Case vbKeyMenu
        KeyCodeName = "Alt"
    Case vbKeyMultiply
        KeyCodeName = "(Num)*"
    Case vbKeyNumpad0
        KeyCodeName = "(Num)0"
    Case vbKeyNumpad1
        KeyCodeName = "(Num)1"
    Case vbKeyNumpad2
        KeyCodeName = "(Num)2"
    Case vbKeyNumpad3
        KeyCodeName = "(Num)3"
    Case vbKeyNumpad4
        KeyCodeName = "(Num)4"
    Case vbKeyNumpad5
        KeyCodeName = "(Num)5"
    Case vbKeyNumpad6
        KeyCodeName = "(Num)6"
    Case vbKeyNumpad7
        KeyCodeName = "(Num)7"
    Case vbKeyNumpad8
        KeyCodeName = "(Num)8"
    Case vbKeyNumpad9
        KeyCodeName = "(Num)9"
    Case vbKeyPageDown
        KeyCodeName = "Page Down"
    Case vbKeyPageUp
        KeyCodeName = "Page Up"
    Case vbKeyPause
        KeyCodeName = "Pause"
    Case vbKeyPrint
        KeyCodeName = "Print"
    Case vbKeyReturn
        KeyCodeName = "Return"
    Case vbKeyRight
        KeyCodeName = "Right"
    Case vbKeySeparator
        KeyCodeName = "(Num)Enter"
    Case vbKeyShift
        KeyCodeName = "Shift"
    Case vbKeySpace
        KeyCodeName = "Space"
    Case vbKeySubtract
        KeyCodeName = "(Num)-"
    Case vbKeyTab
        KeyCodeName = "Tab"
    Case vbKeyUp
        KeyCodeName = "Up"
    Case Asc("0") To Asc("9"), Asc("A") To Asc("Z")
        KeyCodeName = Chr$(KeyCode)
    Case Else
        KeyCodeName = "[Code " & CStr(KeyCode) & "]"
    End Select
        
    KeyControl.Text = KeyCodeName
    KeyControl.Tag = KeyCode
End Sub
