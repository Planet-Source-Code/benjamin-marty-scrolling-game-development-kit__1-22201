VERSION 5.00
Begin VB.Form frmTSDisplay 
   Caption         =   "View Tileset"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   HelpContextID   =   119
   Icon            =   "TSDisp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   50
      Left            =   0
      SmallChange     =   5
      TabIndex        =   1
      Top             =   2880
      Width           =   4455
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2895
      LargeChange     =   50
      Left            =   4440
      SmallChange     =   5
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmTSDisplay"
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
' File: TSDisp.frm - Quick-Display Tileset Window
'
'======================================================================

Option Explicit

Public TSD As TileSetDef

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "ViewTileset", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6)), CLng(Mid$(WndPos, 15, 6)), CLng(Right$(WndPos, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Paint()
    If Not TSD.IsLoaded Then
        Exit Sub
    End If
    If Me.ScaleHeight - HScroll.Height > 0 Then
        Me.PaintPicture TSD.Image, 0, 0, _
            Me.ScaleWidth - VScroll.Width, _
            Me.ScaleHeight - HScroll.Height, _
            HScroll.Value, VScroll.Value, _
            Me.ScaleWidth - VScroll.Width, _
            Me.ScaleHeight - HScroll.Height
    End If
End Sub

Private Sub Form_Resize()
    If Not TSD.IsLoaded Then
        Unload Me
        Exit Sub
    End If
    VScroll.Left = Me.ScaleWidth - VScroll.Width
    HScroll.Top = Me.ScaleHeight - HScroll.Height
    If Me.ScaleHeight - HScroll.Height > 0 Then VScroll.Height = Me.ScaleHeight - HScroll.Height
    HScroll.Width = Me.ScaleWidth - VScroll.Width
    If Me.ScaleY(TSD.Image.Height, vbHimetric, Me.ScaleMode) - Me.ScaleHeight + HScroll.Height > VScroll.Min Then
        VScroll.Max = Me.ScaleY(TSD.Image.Height, vbHimetric, Me.ScaleMode) - Me.ScaleHeight + HScroll.Height
        VScroll.Enabled = True
    Else
        VScroll.Enabled = False
    End If
    If Me.ScaleX(TSD.Image.Width, vbHimetric, Me.ScaleMode) - Me.ScaleWidth + VScroll.Width > HScroll.Min Then
        HScroll.Max = Me.ScaleX(TSD.Image.Width, vbHimetric, Me.ScaleMode) - Me.ScaleWidth + VScroll.Width
        HScroll.Enabled = True
    Else
        HScroll.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "ViewTileset", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000") & "," & Format$(Me.Width, " 00000;-00000") & "," & Format$(Me.Height, " 00000;-00000")
End Sub

Private Sub HScroll_Change()
    Me.Refresh
End Sub

Private Sub HScroll_Scroll()
    Me.Refresh
End Sub

Private Sub VScroll_Change()
    Me.Refresh
End Sub

Private Sub VScroll_Scroll()
    Me.Refresh
End Sub
