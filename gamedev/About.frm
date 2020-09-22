VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About GameDev"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblGameDevHome 
      BackStyle       =   0  'Transparent
      Caption         =   "http://gamedev.sourceforge.net/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "About.frx":0CFA
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label lblBMDXCtls 
      BackStyle       =   0  'Transparent
      Caption         =   "Uses BMDXCtls DirectX library written by Benjamin Marty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "About.frx":0E4C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label lblScrHost 
      BackStyle       =   0  'Transparent
      Caption         =   "Uses ScrHost ActiveX Script Hosting library by Benjamin Marty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label lblGPL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.fsf.org/copyleft/gpl.html"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "About.frx":0F9E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblDist 
      BackStyle       =   0  'Transparent
      Caption         =   "Distributed under GPL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label lblMSAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Uses Microsoft® ActiveX® Scripting technology"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "About.frx":10F0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   960
      X2              =   5400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   120
      Picture         =   "About.frx":1242
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "BlueMonkMN@email.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "About.frx":3684
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Designed and developed by Benjamin Marty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000,2001 Benjamin Marty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrolling Game Development Kit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
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
' File: About.frm - About dialog
'
'======================================================================

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = ""
End Sub

Private Sub lblBMDXCtls_Click()
    Dim TmpDisp As New BMDXDisplay
    
    TmpDisp.ShowAbout
    Set TmpDisp = Nothing
End Sub

Private Sub lblBMDXCtls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = "(Show about dialog for BMDXCtls)"
End Sub

Private Sub lblGameDevHome_Click()
    StartURL lblGameDevHome.Caption
End Sub

Private Sub lblGameDevHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = lblGameDevHome.Caption
End Sub

Private Sub lblGPL_Click()
    StartURL lblGPL.Caption
End Sub

Private Sub lblGPL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = lblGPL.Caption
End Sub

Private Sub lblMSAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = "http://msdn.microsoft.com/scripting/"
End Sub

Private Sub lblEmail_Click()
    StartURL "mailto:BlueMonkMN@email.com"
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = "mailto:BlueMonkMN@email.com"
End Sub

Private Sub lblMSAbout_Click()
    StartURL "http://msdn.microsoft.com/scripting/"
End Sub

Private Sub StartURL(strURL As String)
    On Error Resume Next
    Shell "Start """ & strURL & """"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "Explorer """ & strURL & """"
    End If
    If Err.Number <> 0 Then
        If MsgBox("Can't figure out how to navigate on this OS.  Copy the URL to the clipboard?", vbExclamation + vbYesNo) = vbYes Then
            Clipboard.SetText strURL
        End If
    End If
End Sub

