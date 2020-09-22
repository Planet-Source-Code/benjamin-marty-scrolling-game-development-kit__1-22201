VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4695
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      Begin VB.Timer tmrUnload 
         Interval        =   100
         Left            =   6840
         Top             =   4200
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2000,2001 Benjamin Marty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblLicence 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Distributed under GPL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   1
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         TabIndex        =   3
         Top             =   2700
         Width           =   2445
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scrolling Game Development Kit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   6900
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ben Marty's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5235
         TabIndex        =   4
         Top             =   720
         Width           =   2025
      End
      Begin VB.Image imgLogo 
         Height          =   4800
         Left            =   0
         Picture         =   "Splash.frx":000C
         Top             =   0
         Width           =   3600
      End
   End
End
Attribute VB_Name = "frmSplash"
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
' File: Splash.frm - Splash Screen
'
'======================================================================

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Prj.bSplashShowing = False
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    Prj.bSplashShowing = False
End Sub

Private Sub imgLogo_Click()
    Unload Me
    Prj.bSplashShowing = False
End Sub

Private Sub lblProductName_Click()
    Unload Me
    Prj.bSplashShowing = False
End Sub

Private Sub tmrUnload_Timer()
    Static Timeout As Integer
    Timeout = Timeout + 1
    If Timeout >= 30 Then
        Unload Me
        Prj.bSplashShowing = False
    Else
        Me.ZOrder
    End If
End Sub
