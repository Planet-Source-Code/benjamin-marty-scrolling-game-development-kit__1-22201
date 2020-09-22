VERSION 5.00
Begin VB.Form frmTileImport 
   Caption         =   "Import Tile"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   HelpContextID   =   118
   Icon            =   "TileImp.frx":0000
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
Attribute VB_Name = "frmTileImport"
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
' File: TileImp.frm - Import Tile Graphics Selection Window
'
'======================================================================

Option Explicit

Dim SrcImage As StdPicture
Dim DestImage As StdPicture
Dim TileWidth As Integer
Dim TileHeight As Integer
Dim Px As Integer, Py As Integer
Dim PL As Long, PT As Long

Private Sub Form_Initialize()
    Px = -1000
    Py = -1000
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "TileImport", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6)), CLng(Mid$(WndPos, 15, 6)), CLng(Right$(WndPos, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rcSel As RECT
    
    If Px >= 0 And Py >= 0 Then
        With rcSel
            .Left = Px
            .Top = Py
            .Right = .Left + TileWidth
            .Bottom = .Top + TileHeight
        End With
        
        DrawFocusRect Me.hDC, rcSel
    End If
    Set DestImage = CapturePicture(Me.hDC, Px, Py, TileWidth, TileHeight)
    Me.Hide
End Sub

Public Function ImportTile(Source As StdPicture, ImportTileWidth As Integer, ImportTileHeight As Integer) As StdPicture
    If Int(Me.ScaleX(Source.Width, vbHimetric, vbPixels) + 0.5) < ImportTileWidth Then
        Err.Raise vbObjectError, , "Image is not wide enough"
    ElseIf Int(Me.ScaleY(Source.Height, vbHimetric, vbPixels) + 0.5) < ImportTileHeight Then
        Err.Raise vbObjectError, , "Image is not tall enough"
    End If
    
    TileWidth = ImportTileWidth
    TileHeight = ImportTileHeight
    Set SrcImage = Source
    Me.Show 1
    Set ImportTile = DestImage
    Unload Me
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rcSel As RECT
    
    If TileWidth <= 0 Then Exit Sub
    If TileHeight <= 0 Then Exit Sub
    
    If Px >= -1000 And Py >= -1000 Then
        If PL <> Left Or PT <> Top Then
            Me.Refresh
        Else
            With rcSel
                .Left = Px
                .Top = Py
                .Right = .Left + TileWidth
                .Bottom = .Top + TileHeight
            End With
            
            DrawFocusRect Me.hDC, rcSel
        End If
    End If
    
    With rcSel
        .Left = X - TileWidth \ 2
        .Top = Y - TileHeight \ 2
        If .Left + HScroll.Value < 0 Then .Left = -HScroll.Value
        If .Top + VScroll.Value < 0 Then .Top = -VScroll.Value
        If .Left + HScroll.Value + TileWidth > Int(Me.ScaleX(SrcImage.Width, vbHimetric, vbPixels) + 0.5) Then
            .Left = Int(Me.ScaleX(SrcImage.Width, vbHimetric, vbPixels) + 0.5) - TileWidth - HScroll.Value
        End If
        If .Top + VScroll.Value + TileHeight > Int(Me.ScaleY(SrcImage.Height, vbHimetric, vbPixels) + 0.5) Then
            .Top = Int(Me.ScaleY(SrcImage.Height, vbHimetric, vbPixels) + 0.5) - TileHeight - VScroll.Value
        End If
        .Right = .Left + TileWidth
        .Bottom = .Top + TileHeight
        Px = .Left
        Py = .Top
    End With
    PL = Left
    PT = Top
    
    DrawFocusRect Me.hDC, rcSel
    Me.Caption = "Import Tile - " & Px & "," & Py
End Sub

Private Sub Form_Paint()
    If SrcImage Is Nothing Then
        Exit Sub
    End If
    If Me.ScaleHeight - HScroll.Height > 0 Then
        Me.PaintPicture SrcImage, 0, 0, _
            Me.ScaleWidth - VScroll.Width, _
            Me.ScaleHeight - HScroll.Height, _
            HScroll.Value, VScroll.Value, _
            Me.ScaleWidth - VScroll.Width, _
            Me.ScaleHeight - HScroll.Height
    End If
    Px = -1000
    Py = -1000
End Sub

Private Sub Form_Resize()
    If SrcImage Is Nothing Then
        Exit Sub
    End If
    VScroll.Left = Me.ScaleWidth - VScroll.Width
    HScroll.Top = Me.ScaleHeight - HScroll.Height
    If Me.ScaleHeight - HScroll.Height > 0 Then VScroll.Height = Me.ScaleHeight - HScroll.Height
    HScroll.Width = Me.ScaleWidth - VScroll.Width
    If Int(Me.ScaleY(SrcImage.Height, vbHimetric, Me.ScaleMode) + 0.5) - Me.ScaleHeight + HScroll.Height > VScroll.Min Then
        VScroll.Max = Int(Me.ScaleY(SrcImage.Height, vbHimetric, Me.ScaleMode) + 0.5) - Me.ScaleHeight + HScroll.Height
        VScroll.Enabled = True
    Else
        VScroll.Enabled = False
    End If
    If Int(Me.ScaleX(SrcImage.Width, vbHimetric, Me.ScaleMode) + 0.5) - Me.ScaleWidth + VScroll.Width > HScroll.Min Then
        HScroll.Max = Int(Me.ScaleX(SrcImage.Width, vbHimetric, Me.ScaleMode) + 0.5) - Me.ScaleWidth + VScroll.Width
        HScroll.Enabled = True
    Else
        HScroll.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "TileImport", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000") & "," & Format$(Me.Width, " 00000;-00000") & "," & Format$(Me.Height, " 00000;-00000")
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
