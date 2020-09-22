VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTileAnim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Animation"
   ClientHeight    =   6255
   ClientLeft      =   1020
   ClientTop       =   705
   ClientWidth     =   6510
   HelpContextID   =   117
   Icon            =   "TileAnim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDefineAnims 
      Caption         =   "Define Tile Animation (Drag && Drop)"
      Enabled         =   0   'False
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   6255
      Begin VB.Timer tmrPreview 
         Interval        =   20
         Left            =   4800
         Top             =   3120
      End
      Begin MSComCtl2.UpDown udDelay 
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtDelay"
         BuddyDispid     =   196611
         OrigLeft        =   4200
         OrigTop         =   3345
         OrigRight       =   4440
         OrigBottom      =   3645
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         Top             =   3360
         Width           =   615
      End
      Begin VB.PictureBox picPreview 
         Height          =   975
         Left            =   4800
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.PictureBox picTile 
         Height          =   975
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.HScrollBar hscrollFrames 
         Height          =   255
         LargeChange     =   60
         Left            =   1440
         SmallChange     =   5
         TabIndex        =   18
         Top             =   3000
         Width           =   3135
      End
      Begin VB.PictureBox picFrames 
         Height          =   975
         Left            =   1440
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   205
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3135
      End
      Begin VB.PictureBox picTileset 
         Height          =   975
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   381
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   5775
      End
      Begin VB.VScrollBar vscrollTileset 
         Height          =   975
         LargeChange     =   30
         Left            =   5880
         SmallChange     =   5
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Set delay for selected frame:"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblPreview 
         Caption         =   "Preview:"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTileToAnim 
         Caption         =   "Tile to Animate:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblAnimFrames 
         Caption         =   "Animation Frames"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lblTileset 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiles in Layer's Tileset:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Frame fraManageAnims 
      Caption         =   "Manage Tile Animations:"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.ListBox lstLayers 
         Height          =   645
         Left            =   120
         TabIndex        =   4
         Top             =   1455
         Width           =   2295
      End
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename Anim."
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Anim."
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Anim."
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox lstAnimDefs 
         Height          =   1620
         Left            =   2520
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox lstMaps 
         Height          =   645
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblLayers 
         Caption         =   "Layers in Map:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1215
         Width           =   2295
      End
      Begin VB.Label lblAnimDefs 
         Caption         =   "Animations Defined for Layer:"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblMaps 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Maps:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmTileAnim"
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
' File: TileAnim.frm - Tile Animation Definition Dialog
'
'======================================================================

Option Explicit

Dim EditAnim As AnimDef
Dim Highlighted As New TileGroup
Dim DragPic As StdPicture
Dim StartDragPt As POINTAPI
Dim DragState As Integer
Dim CFTiles As Long
Dim CFFrames As Long
Dim pInsIdx As Integer
Dim FrameHighlight As New TileGroup
Dim bPreviewRun As Boolean

Sub UpdateMaps()
    Dim I As Integer
    
    lstMaps.Clear
    For I = 0 To Prj.MapCount - 1
        lstMaps.AddItem Prj.Maps(I).Name
    Next I
    
End Sub

Sub UpdateLayers()
    Dim I As Integer
    
    lstLayers.Clear
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    
    With Prj.Maps(lstMaps.List(lstMaps.ListIndex))
        For I = 0 To .LayerCount - 1
            lstLayers.AddItem .MapLayer(I).Name
        Next
    End With
    
End Sub

Sub UpdateAnims()
    Dim I As Integer
    Dim MapName As String
    Dim LayerName As String
    
    lstAnimDefs.Clear
    Set EditAnim = Nothing
    UpdateEditor
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstLayers.ListIndex < 0 Then Exit Sub
    
    MapName = lstMaps.List(lstMaps.ListIndex)
    LayerName = lstLayers.List(lstLayers.ListIndex)
    
    For I = 0 To Prj.AnimDefCount - 1
        If Prj.AnimDefs(I).MapName = MapName And Prj.AnimDefs(I).LayerName = LayerName Then
            lstAnimDefs.AddItem Prj.AnimDefs(I).Name
        End If
    Next I
    
End Sub

Sub CheckAnimAssoc()
    Dim I As Integer
    Dim bFound As Boolean
    Dim bRepeat As Boolean
    
    Do
        bRepeat = False
        For I = 0 To Prj.AnimDefCount - 1
            With Prj.AnimDefs(I)
                If Prj.MapExists(.MapName) Then
                    If Prj.Maps(.MapName).LayerExists(.LayerName) Then
                        bFound = True
                    Else
                        bFound = False
                    End If
                Else
                    bFound = False
                End If
                If Not bFound Then
                    If MsgBox("Tile Animation Definition """ & .Name & """ is associated with a map or layer that no longer exists.  Delete this Animation Definition?", vbYesNo) = vbYes Then
                        Prj.RemoveAnim .Name
                        bRepeat = True
                        Exit For
                    End If
                End If
            End With
        Next I
    Loop While bRepeat
    
End Sub

Private Sub cmdDelete_Click()
    If lstAnimDefs.ListIndex < 0 Then
        MsgBox "Please select a map layer and an animation before selecting this command."
        Exit Sub
    End If
    Prj.RemoveAnim lstAnimDefs.List(lstAnimDefs.ListIndex)
    UpdateAnims
End Sub

Private Sub cmdNew_Click()
    Dim A As New AnimDef
    
    If lstMaps.ListIndex < 0 Or lstLayers.ListIndex < 0 Then
        MsgBox "Please select a map layer before selecting this command."
        Exit Sub
    End If
    
    A.MapName = lstMaps.List(lstMaps.ListIndex)
    A.LayerName = lstLayers.List(lstLayers.ListIndex)
    A.Name = InputBox$("Enter a name for the tile animation:", "Create Tile Animation Definition")
    If Len(A.Name) > 0 Then
        Prj.AddAnim A
        UpdateAnims
        lstAnimDefs.ListIndex = lstAnimDefs.NewIndex
    End If
    
End Sub

Public Sub UpdateEditor()
    If EditAnim Is Nothing Then
        fraDefineAnims.Enabled = False
    Else
        fraDefineAnims.Enabled = True
        Highlighted.ClearAll
        FrameHighlight.ClearAll
        PaintTileset
        PaintFrames
        PaintBaseTile
    End If
    
End Sub

Private Sub cmdRename_Click()
    Dim NewVal As String
    
    If lstAnimDefs.ListIndex < 0 Then
        MsgBox "Please select a map layer and an animation before selecting this command."
        Exit Sub
    End If
    
    NewVal = InputBox$("Enter new name", "Rename Animation Definition")
    If Len(NewVal) > 0 Then
        Prj.AnimDefs(lstAnimDefs.List(lstAnimDefs.ListIndex)).Name = NewVal
    End If
    UpdateAnims
End Sub

Private Sub Form_Initialize()
    CFTiles = RegisterClipboardFormat("TileGroup")
    CFFrames = RegisterClipboardFormat("FrameGroup")
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "TileAnim", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    CheckAnimAssoc
    UpdateMaps
    UpdateLayers
    UpdateAnims
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bPreviewRun = False
    SaveSetting "GameDev", "Windows", "TileAnim", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub hscrollFrames_Change()
    PaintFrames
End Sub

Private Sub hscrollFrames_Scroll()
    PaintFrames
End Sub

Private Sub lstAnimDefs_Click()
    If lstAnimDefs.ListIndex >= 0 Then
        Set EditAnim = Prj.AnimDefs(lstAnimDefs.List(lstAnimDefs.ListIndex))
        UpdateEditor
    End If
End Sub

Private Sub lstLayers_Click()
    UpdateAnims
End Sub

Private Sub lstMaps_Click()
    UpdateLayers
    UpdateAnims
End Sub

Function GetTSDef() As TileSetDef
    Set GetTSDef = Prj.Maps(EditAnim.MapName).MapLayer(EditAnim.LayerName).TSDef
End Function

Sub PaintTileset()
    Dim TSCols As Integer
    Dim TSRows As Integer
    Dim I As Integer
    Dim rcTile As RECT
    Dim YMax As Integer
    
    If EditAnim Is Nothing Then Exit Sub

    With GetTSDef
        On Error Resume Next
        If Not .IsLoaded Then .Load
        If Not .IsLoaded Then Exit Sub
        On Error GoTo 0
        
        TSCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight
    End With
    
    picTileset.Cls
    For I = 0 To TSCols * TSRows - 1
        rcTile = GetTilesetTileRect(I)
        If rcTile.Top > YMax Then YMax = rcTile.Top
        rcTile.Top = rcTile.Top - vscrollTileset.Value
        rcTile.Bottom = rcTile.Bottom - vscrollTileset.Value
        If rcTile.Bottom > 0 And rcTile.Top <= picTileset.ScaleHeight Then
            If Highlighted.IsMember(I) Then
                picTileset.Line (rcTile.Left - 2, rcTile.Top - 2)-(rcTile.Right + 2, rcTile.Bottom + 2), vbBlue, BF
            End If
            picTileset.PaintPicture ExtractLocalTile(I, Highlighted.IsMember(I)), rcTile.Left, rcTile.Top
        End If
    Next
    
    vscrollTileset.Max = YMax
    
End Sub

Private Function ExtractLocalTile(ByVal Index As Integer, Optional bHighlight As Boolean = False) As StdPicture
    Dim TSD As TileSetDef
    Dim TSCols As Integer
    Dim TSRows As Integer
    
    With GetTSDef
        If .Image Is Nothing Then
            .Load
        End If
        If .Image Is Nothing Then Exit Function
        TSCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight
        If Index < TSRows * TSCols Then
            Set ExtractLocalTile = ExtractTile(.Image, .TileWidth * (Index Mod TSCols), .TileHeight * (Index \ TSCols), .TileWidth, .TileHeight, bHighlight)
        Else
            MsgBox "Tile index out of bounds", vbExclamation, "ExtractLocalTile"
        End If
    End With
    
End Function

Private Function GetTilesetTileRect(Index) As RECT
    Dim FitCols As Integer
        
    With GetTSDef
        FitCols = (Me.picTileset.ScaleWidth) \ (.TileWidth + 6)
        GetTilesetTileRect.Left = (Index Mod FitCols) * (.TileWidth + 6) + 3
        GetTilesetTileRect.Top = (Index \ FitCols) * (.TileHeight + 6) + 3
        GetTilesetTileRect.Right = GetTilesetTileRect.Left + .TileWidth - 1
        GetTilesetTileRect.Bottom = GetTilesetTileRect.Top + .TileHeight - 1
    End With
    
End Function

Private Function GetXYTileSetTile(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim TSCols As Integer
    Dim TSRows As Integer
    Dim I As Integer
    Dim rcTile As RECT
    
    GetXYTileSetTile = -1
    
    If EditAnim Is Nothing Then Exit Function

    With GetTSDef
        On Error Resume Next
        If Not .IsLoaded Then .Load
        If Not .IsLoaded Then Exit Function
        On Error GoTo 0
        
        TSCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight
    End With
    
    For I = 0 To TSCols * TSRows - 1
        rcTile = GetTilesetTileRect(I)
        rcTile.Top = rcTile.Top - vscrollTileset.Value
        rcTile.Bottom = rcTile.Bottom - vscrollTileset.Value
        If X >= rcTile.Left And X <= rcTile.Right And Y >= rcTile.Top And Y <= rcTile.Bottom Then
            GetXYTileSetTile = I
            Exit Function
        End If
    Next
        
End Function

Private Sub picFrames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 1
End Sub

Private Sub picFrames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragState = 1 Then
        If Abs(StartDragPt.X - X) > 3 Or Abs(StartDragPt.Y - Y) > 3 Then
            DragState = 2
            picFrames.OLEDrag
        End If
    End If
End Sub

Private Sub picFrames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Idx As Integer
    
    Idx = GetXYFrameTile(X, Y)
    
    If DragState = 1 And Idx >= 0 Then
        If (Shift And vbCtrlMask) = 0 Then
            SelectSingleFrame Idx
        End If
    End If
    
    DragState = 0

End Sub

Private Sub picFrames_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    Dim InsIdx As Integer
    Dim TempAnim As AnimDef
    Dim Delta As Integer
    
    If Data.GetFormat(CInt("&H" & Hex$(CFFrames))) Then
        Set TempAnim = EditAnim.Clone
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFFrames)))
        TileCount = DropTiles(LBound(DropTiles))
        InsIdx = GetFrameXYInsertIndex(X, Y)
        Effect = vbDropEffectCopy
        If Shift And vbShiftMask Then
            Effect = vbDropEffectMove
            For I = 0 To EditAnim.FrameCount - 1
                If FrameHighlight.IsMember(I) Then
                    If InsIdx > I + Delta Then InsIdx = InsIdx - 1
                    EditAnim.RemoveFrame I + Delta
                    Delta = Delta - 1
                End If
            Next
        End If
        For I = LBound(DropTiles) + TileCount To LBound(DropTiles) + 1 Step -1
            EditAnim.InsertFrame InsIdx, TempAnim.FrameValue(DropTiles(I)), Val(txtDelay.Text)
        Next
        Set TempAnim = Nothing
    ElseIf Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        InsIdx = GetFrameXYInsertIndex(X, Y)
        For I = LBound(DropTiles) + TileCount To LBound(DropTiles) + 1 Step -1
            EditAnim.InsertFrame InsIdx, DropTiles(I), Val(txtDelay.Text)
        Next
    End If

    Prj.IsDirty = True
    
    DragState = 0
    pInsIdx = -1
    
    FrameHighlight.ClearAll
    PaintFrames
    
End Sub

Private Sub picFrames_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim InsIdx As Integer
    
    If EditAnim Is Nothing Then Exit Sub
    
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        If Shift And vbCtrlMask Then
            Effect = vbDropEffectCopy
        ElseIf Shift And vbShiftMask Then
            Effect = vbDropEffectMove
        Else
            Effect = vbDropEffectMove Or vbDropEffectCopy
        End If
        InsIdx = GetFrameXYInsertIndex(X, Y)
        If State = vbOver Then
            With GetTSDef
                If (InsIdx - 1) * (.TileWidth + 6) < hscrollFrames.Value Then
                    If hscrollFrames.Value - .TileWidth - 6 > hscrollFrames.Min Then
                        hscrollFrames.Value = hscrollFrames.Value - .TileWidth - 6
                        pInsIdx = -1
                    Else
                        If hscrollFrames.Value <> hscrollFrames.Min Then
                            hscrollFrames.Value = hscrollFrames.Min
                            pInsIdx = -1
                        End If
                    End If
                End If
                If (InsIdx + 1) * (.TileWidth + 6) > hscrollFrames.Value + picFrames.ScaleWidth Then
                    If hscrollFrames.Value + .TileWidth + 6 < hscrollFrames.Max Then
                        hscrollFrames.Value = hscrollFrames.Value + .TileWidth + 6
                        pInsIdx = -1
                    Else
                        If hscrollFrames.Value <> hscrollFrames.Max Then
                            hscrollFrames.Value = hscrollFrames.Max
                            pInsIdx = -1
                        End If
                    End If
                End If
                If pInsIdx <> InsIdx Then
                    PaintFrames
                    picFrames.Line (InsIdx * (.TileWidth + 6) - 1 - hscrollFrames.Value, 0)-(InsIdx * (.TileWidth + 6) - hscrollFrames.Value, picFrames.ScaleHeight - 1), , B
                    pInsIdx = InsIdx
                End If
            End With
        Else
            PaintFrames
        End If
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Function GetFrameXYInsertIndex(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim I As Integer
    Dim rcFrame As RECT

    If EditAnim Is Nothing Then Exit Function
    
    For I = 0 To EditAnim.FrameCount - 1
        rcFrame = GetFrameRect(I)
        rcFrame.Left = rcFrame.Left - hscrollFrames.Value
        rcFrame.Right = rcFrame.Right - hscrollFrames.Value
        If X < (rcFrame.Left + rcFrame.Right) / 2 Then
            GetFrameXYInsertIndex = I
            Exit Function
        End If
    Next
    
    GetFrameXYInsertIndex = I
    
End Function

Private Sub picFrames_OLESetData(Data As DataObject, DataFormat As Integer)
    Dim DragTiles() As Byte
    Dim V As Variant
    Dim I As Integer
    
    If DataFormat = vbCFDIB Then
        Data.SetData DragPic, vbCFDIB
    ElseIf Hex$(DataFormat) = Hex$(CFFrames) Then
        V = FrameHighlight.GetArray
        ReDim DragTiles(LBound(V) To UBound(V) + 1)
        DragTiles(LBound(V)) = UBound(V) - LBound(V) + 1 ' Tile count
        For I = LBound(V) To UBound(V)
            DragTiles(I + 1) = V(I)
        Next
        Data.SetData DragTiles, CInt("&H" & Hex$(CFFrames))
    ElseIf Hex$(DataFormat) = Hex$(CFTiles) Then
        V = FrameHighlight.GetArray
        ReDim DragTiles(LBound(V) To UBound(V) + 1)
        DragTiles(LBound(V)) = UBound(V) - LBound(V) + 1 ' Tile count
        For I = LBound(V) To UBound(V)
            DragTiles(I + 1) = EditAnim.FrameValue(V(I))
        Next
        Data.SetData DragTiles, CInt("&H" & Hex$(CFTiles))
    End If
End Sub

Private Sub picFrames_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Data.SetData Format:=vbCFDIB
    Data.SetData Format:=CInt("&H" & Hex$(CFFrames))
    Data.SetData Format:=CInt("&H" & Hex$(CFTiles))
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
End Sub

Private Sub picFrames_Paint()
    PaintFrames
End Sub

Private Sub picTile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectCopy
        EditAnim.BaseTile = Data.GetData(CInt("&H" & Hex$(CFTiles)))(2)
    Else
        Effect = vbDropEffectNone
    End If
    
    Prj.IsDirty = True
    
    DragState = 0
    PaintBaseTile
    
End Sub

Private Sub picTile_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub picTile_Paint()
    PaintBaseTile
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 0
End Sub

Private Sub HandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Grp As Integer)
    Dim Idx As Integer
    Dim H As TileGroup
    
    If Button = 1 Then
        If Grp = 0 Then
            Idx = GetXYTileSetTile(X, Y)
            Set H = Highlighted
        ElseIf Grp = 1 Then
            Idx = GetXYFrameTile(X, Y)
            Set H = FrameHighlight
        End If
        
        If Idx >= 0 Then
            If Shift And vbCtrlMask Then
                If H.IsMember(Idx) Then
                    H.ClearMember Idx
                Else
                    H.SetMember Idx
                    Set DragPic = ExtractLocalTile(Idx)
                    DragState = 1
                    StartDragPt.X = X
                    StartDragPt.Y = Y
                End If
                PaintTileset
                PaintFrames
            Else
                If Not H.IsMember(Idx) Then
                    If Grp = 0 Then
                        SelectSingle Idx
                    Else
                        SelectSingleFrame Idx
                    End If
                End If
                If Grp = 0 Then
                    Set DragPic = ExtractLocalTile(Idx)
                Else
                    Set DragPic = ExtractLocalTile(EditAnim.FrameValue(Idx))
                End If
                DragState = 1
                StartDragPt.X = X
                StartDragPt.Y = Y
            End If
        Else
            If (Shift And vbCtrlMask) = 0 Then
                If Not H.IsEmpty Then
                    H.ClearAll
                    If Grp = 0 Then
                        PaintTileset
                    Else
                        PaintFrames
                    End If
                End If
            End If
            DragState = 0
        End If
    End If
    
End Sub

Sub SelectSingle(Idx As Integer)
    Dim V As Variant
    Dim bUpdate As Boolean
    
    V = Highlighted.GetArray
    If IsEmpty(V) Then
        bUpdate = True
    Else
        If UBound(V) - LBound(V) > 0 Then bUpdate = True
        If V(LBound(V)) <> Idx Then bUpdate = True
    End If
    If bUpdate Then
        Highlighted.ClearAll
        Highlighted.SetMember Idx
        PaintTileset
    End If
End Sub

Sub SelectSingleFrame(Idx As Integer)
    Dim V As Variant
    Dim bUpdate As Boolean
    
    V = FrameHighlight.GetArray
    If IsEmpty(V) Then
        bUpdate = True
    Else
        If UBound(V) - LBound(V) > 0 Then bUpdate = True
        If V(LBound(V)) <> Idx Then bUpdate = True
    End If
    If bUpdate Then
        FrameHighlight.ClearAll
        FrameHighlight.SetMember Idx
        PaintFrames
    End If
    txtDelay.Text = CStr(EditAnim.FrameDelay(Idx))
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragState = 1 Then
        If Abs(StartDragPt.X - X) > 3 Or Abs(StartDragPt.Y - Y) > 3 Then
            picTileset.OLEDrag
            DragState = 2
        End If
    End If
End Sub

Private Sub picTileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Idx As Integer
    
    Idx = GetXYTileSetTile(X, Y)
    
    If DragState = 1 And Idx >= 0 Then
        If (Shift And vbCtrlMask) = 0 Then
            SelectSingle Idx
        End If
    End If
    
    DragState = 0
    
End Sub

Sub PaintFrames()
    Dim I As Integer
    Dim rcTile As RECT
    Dim XMax As Integer
    
    If EditAnim Is Nothing Then Exit Sub
    
    pInsIdx = -1 ' Force repaint of insertion point
    
    picFrames.Cls
    For I = 0 To EditAnim.FrameCount - 1
        rcTile = GetFrameRect(I)
        If rcTile.Right - picFrames.ScaleWidth + 6 > XMax Then XMax = rcTile.Right - picFrames.ScaleWidth + 6
        rcTile.Left = rcTile.Left - hscrollFrames.Value
        rcTile.Right = rcTile.Right - hscrollFrames.Value
        If rcTile.Right > 0 And rcTile.Left <= picFrames.ScaleWidth Then
            If FrameHighlight.IsMember(I) Then
                picFrames.Line (rcTile.Left - 2, rcTile.Top - 2)-(rcTile.Right + 2, rcTile.Bottom + 2), vbBlue, BF
            End If
            picFrames.PaintPicture ExtractLocalTile(EditAnim.FrameValue(I), FrameHighlight.IsMember(I)), rcTile.Left, rcTile.Top
        End If
    Next
    
    hscrollFrames.Max = XMax
    tmrPreview.Enabled = True
    
End Sub

Private Function GetFrameRect(Index As Integer) As RECT
        
    With GetTSDef
        GetFrameRect.Left = Index * (.TileWidth + 6) + 3
        GetFrameRect.Top = (picFrames.ScaleHeight - .TileHeight) / 2
        GetFrameRect.Right = GetFrameRect.Left + .TileWidth - 1
        GetFrameRect.Bottom = GetFrameRect.Top + .TileHeight - 1
    End With
    
End Function

Private Function GetXYFrameTile(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim I As Integer
    Dim rcTile As RECT
    Dim XMax As Integer
    
    GetXYFrameTile = -1
    
    If EditAnim Is Nothing Then Exit Function
    
    For I = 0 To EditAnim.FrameCount - 1
        rcTile = GetFrameRect(I)
        rcTile.Left = rcTile.Left - hscrollFrames.Value
        rcTile.Right = rcTile.Right - hscrollFrames.Value
        If X >= rcTile.Left And X <= rcTile.Right And Y >= rcTile.Top And Y <= rcTile.Bottom Then
            GetXYFrameTile = I
            Exit Function
        End If
    Next

End Function

Private Sub PaintBaseTile()
    
    If EditAnim Is Nothing Then Exit Sub
    
    With GetTSDef
        picTile.PaintPicture ExtractLocalTile(EditAnim.BaseTile), (picTile.ScaleWidth - .TileWidth) / 2, (picTile.ScaleHeight - .TileHeight) / 2
    End With
    
End Sub

Private Sub picTileset_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    Dim Delta As Integer
    
    If Data.GetFormat(CInt("&H" & Hex$(CFFrames))) Then
        Effect = vbDropEffectMove
        For I = 0 To EditAnim.FrameCount - 1
            If FrameHighlight.IsMember(I) Then
                EditAnim.RemoveFrame I + Delta
                Delta = Delta - 1
            End If
        Next
    End If

    Prj.IsDirty = True
    
    DragState = 0
    pInsIdx = -1
    
    FrameHighlight.ClearAll
    PaintFrames
    
End Sub

Private Sub picTileset_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    If Data.GetFormat(CInt("&H" & Hex$(CFFrames))) Then
        Effect = vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub picTileset_OLESetData(Data As DataObject, DataFormat As Integer)
    Dim DragTiles() As Byte
    Dim V As Variant
    Dim I As Integer
    
    If DataFormat = vbCFDIB Then
        Data.SetData DragPic, vbCFDIB
    ElseIf Hex$(DataFormat) = Hex$(CFTiles) Then
        V = Highlighted.GetArray
        ReDim DragTiles(LBound(V) To UBound(V) + 1)
        DragTiles(LBound(V)) = UBound(V) - LBound(V) + 1 ' Tile count
        For I = LBound(V) To UBound(V)
            DragTiles(I + 1) = V(I)
        Next
        Data.SetData DragTiles, CInt("&H" & Hex$(CFTiles))
    End If

End Sub

Private Sub picTileset_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Data.SetData Format:=vbCFDIB
    Data.SetData Format:=CInt("&H" & Hex$(CFTiles))
    AllowedEffects = vbDropEffectCopy
End Sub

Private Sub picTileset_Paint()
    PaintTileset
End Sub

Private Sub tmrPreview_Timer()
    Dim nTimerSpeed As Long
    Dim T As Single
    Dim I As Integer
    
    tmrPreview.Enabled = False
    If bPreviewRun Then Exit Sub
    bPreviewRun = True
    
    T = Timer
    Do
        nTimerSpeed = nTimerSpeed + 1
        DoEvents
    Loop Until Timer - T >= 0.5

    If bPreviewRun = False Then Exit Sub

    nTimerSpeed = nTimerSpeed / 10

    Do While bPreviewRun
    
        bPreviewRun = False
        If EditAnim Is Nothing Then Exit Sub
        If EditAnim.FrameCount = 0 Then Exit Sub
        bPreviewRun = True
        
        With GetTSDef
            picPreview.PaintPicture ExtractLocalTile(EditAnim.CurTile), (picPreview.ScaleWidth - .TileWidth) / 2, (picPreview.ScaleHeight - .TileHeight) / 2
        End With
        EditAnim.Advance
        
        T = 0
        Do
            T = T + 1
            DoEvents
        Loop Until T >= nTimerSpeed
    Loop

End Sub

Private Sub txtDelay_Change()
    Dim V As Variant
    
    V = FrameHighlight.GetArray
    
    If IsEmpty(V) Then
        If txtDelay.Text <> "" Then
            Beep
            txtDelay.Text = ""
        End If
        Exit Sub
    End If
    
    If UBound(V) - LBound(V) > 0 Then
        If txtDelay.Text <> "" Then
            Beep
            txtDelay.Text = ""
        End If
        Exit Sub
    End If
    
    Prj.IsDirty = True
    
    EditAnim.FrameDelay(V(LBound(V))) = Val(txtDelay.Text)
    
End Sub

Private Sub vscrollTileset_Change()
    PaintTileset
End Sub

Private Sub vscrollTileset_Scroll()
    PaintTileset
End Sub
