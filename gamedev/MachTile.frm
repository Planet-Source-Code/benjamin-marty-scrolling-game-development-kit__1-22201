VERSION 5.00
Begin VB.Form frmMatchTile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Tile Matching"
   ClientHeight    =   6615
   ClientLeft      =   1755
   ClientTop       =   345
   ClientWidth     =   5295
   HelpContextID   =   104
   Icon            =   "MachTile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScrollTileSet 
      Height          =   975
      LargeChange     =   30
      Left            =   4920
      SmallChange     =   5
      TabIndex        =   9
      Top             =   5520
      Width           =   255
   End
   Begin VB.PictureBox picTileSet 
      Height          =   975
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   8
      ToolTipText     =   "Drag tiles from here to add them to the current tile matching group, or drag them to here to remove them from the group."
      Top             =   5520
      Width           =   4815
   End
   Begin VB.PictureBox picGroup 
      Height          =   975
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   5
      ToolTipText     =   "Tiles in this box will cause neighboring tiles to match, but will never be placed automatically by the map editor"
      Top             =   4200
      Width           =   4815
   End
   Begin VB.PictureBox picSelected 
      Height          =   975
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   2
      ToolTipText     =   $"MachTile.frx":0442
      Top             =   2880
      Width           =   4815
   End
   Begin VB.VScrollBar VScrollSelected 
      Height          =   975
      LargeChange     =   30
      Left            =   4920
      SmallChange     =   5
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.VScrollBar VScrollGroup 
      Height          =   975
      LargeChange     =   30
      Left            =   4920
      SmallChange     =   5
      TabIndex        =   6
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   14
      Left            =   2640
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   13
      Left            =   2040
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles available in tileset:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label lblSelected 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles in current slot:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label lblGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "Unclassified tiles in this group:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"MachTile.frx":04DC
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   4
      Left            =   2640
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   3
      Left            =   2040
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   2
      Left            =   1320
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   1
      Left            =   720
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   0
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   8
      Left            =   2040
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   9
      Left            =   2640
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   12
      Left            =   1320
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   10
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   6
      Left            =   720
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   11
      Left            =   720
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   7
      Left            =   1320
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgMatchTile 
      Height          =   495
      Index           =   5
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmMatchTile"
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
' File: MachTile.frm - Tile Matching Definition Dialog
'
'======================================================================

Option Explicit

Public EditTileMatch As MatchDef
Public Highlighted As TileGroup
Dim SelGroup As Integer
Dim DragState As Integer
Dim DragPic As StdPicture
Dim CFTiles As Long
Dim StartDragPt As POINTAPI

Private Sub Form_Initialize()
    CFTiles = RegisterClipboardFormat("TileGroup")
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "TileMatching", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Paint()
    Dim I As Integer
    
    For I = imgMatchTile.LBound To imgMatchTile.UBound
        PaintTile I
    Next
End Sub

Sub PaintTile(ByVal Index As Integer)
    Dim X As Integer, C As Byte
    Dim Max As Integer
    Dim L As Integer, T As Integer
    
    With imgMatchTile(Index)
        L = .Left
        T = .Top
        Max = .Width - 1
    End With
    
    If Index = SelGroup Then
        Line (L - 2, T - 2)-(L + Max + 2, T + Max + 2), vbBlue, BF
    End If
    
    If Index = 6 Then
        If Index = SelGroup Then
            Line (L, T)-(L + Max, T + Max), RGB(127, 127, 255), BF
        Else
            Line (L, T)-(L + Max, T + Max), RGB(255, 255, 255), BF
        End If
        Exit Sub
    End If

    If Not EditTileMatch.TileMatches.MatchGroup(Index).IsEmpty Then
        Exit Sub
    End If
    
    For X = 0 To Max
        C = Int(X * 255 / Max)
        If Index = SelGroup Then
            Me.ForeColor = RGB(C * 2 / 3, C * 2 / 3, (C * 2 + 255) / 3)
        Else
            Me.ForeColor = RGB(C, C, C)
        End If
        
        Select Case Index
        Case 0
            Line (L + Max, T + X)-(L + X, T + X)
            Line -(L + X, T + Max)
        Case 1
            Line (L, T + X)-(L + Max, T + X)
        Case 2
            Line (L, T + X)-(L + Max - X, T + X)
            Line -(L + Max - X, T + Max)
        Case 3
            Line (L, T + X)-(L + X, T + X)
            Line -(L + X, T)
        Case 4
            Line (L + Max, T + X)-(L + Max - X, T + X)
            Line -(L + Max - X, T)
        Case 5
            Line (L + X, T)-(L + X, T + Max)
        Case 7
            Line (L + Max - X, T)-(L + Max - X, T + Max)
        Case 8
            Line (L, T + Max - X)-(L + X, T + Max - X)
            Line -(L + X, T + Max)
        Case 9
            Line (L + Max, T + Max - X)-(L + Max - X, T + Max - X)
            Line -(L + Max - X, T + Max)
        Case 10
            Line (L + X, T)-(L + X, T + Max - X)
            Line -(L + Max, T + Max - X)
        Case 11
            Line (L, T + Max - X)-(L + Max, T + Max - X)
        Case 12
            Line (L + Max - X, T)-(L + Max - X, T + Max - X)
            Line -(L, T + Max - X)
        Case 13
            Line (L, T + Max - X)-(L + X, T + Max)
            Line (L + Max - X, T)-(L + Max, T + X)
        Case 14
            Line (L, T + X)-(L + X, T)
            Line (L + Max - X, T + Max)-(L + Max, T + Max - X)
        End Select
    Next
End Sub

Public Sub EditMatches(MD As MatchDef)
    Set EditTileMatch = MD
    Set Highlighted = New TileGroup
    UpdateImages
    Me.Show
End Sub

Public Sub PaintGroup()
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim YMax As Long
    Dim V As Variant
    
    picGroup.Cls
    
    V = EditTileMatch.AllTiles.GetArray
    If IsEmpty(V) Then Exit Sub
    For J = LBound(V) To UBound(V)
        I = V(J)
        If Not EditTileMatch.TileMatches.IsMember(I) Then
            With EditTileMatch.TSDef
                rcDraw = GetTileRect(DrawIndex)
                If rcDraw.Top > YMax Then YMax = rcDraw.Top
                rcDraw.Top = rcDraw.Top - VScrollGroup.Value
                rcDraw.Bottom = rcDraw.Bottom - VScrollGroup.Value
                If rcDraw.Bottom > 0 And rcDraw.Top < picSelected.ScaleHeight Then
                    If Highlighted.IsMember(I) Then
                        picGroup.Line (rcDraw.Left - 2, rcDraw.Top - 2)-(rcDraw.Right + 2, rcDraw.Bottom + 2), vbBlue, BF
                        picGroup.PaintPicture ExtractLocalTile(I, True), rcDraw.Left, rcDraw.Top
                    Else
                        picGroup.PaintPicture ExtractLocalTile(I), rcDraw.Left, rcDraw.Top
                    End If
                End If
                DrawIndex = DrawIndex + 1
            End With
        End If
    Next
    
    VScrollGroup.Max = YMax
    
End Sub

Public Sub PaintSelected()
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim YMax As Long
    Dim V As Variant
    
    picSelected.Cls
    
    If SelGroup < 0 Then Exit Sub
    
    V = EditTileMatch.AllTiles.GetArray
    If IsEmpty(V) Then Exit Sub
    For J = LBound(V) To UBound(V)
        I = V(J)
        If EditTileMatch.TileMatches.MatchGroup(SelGroup).IsMember(I) Then
            With EditTileMatch.TSDef
                rcDraw = GetTileRect(DrawIndex)
                If rcDraw.Top > YMax Then YMax = rcDraw.Top
                rcDraw.Top = rcDraw.Top - VScrollSelected.Value
                rcDraw.Bottom = rcDraw.Bottom - VScrollSelected.Value
                If rcDraw.Bottom > 0 And rcDraw.Top < picSelected.ScaleHeight Then
                    If Highlighted.IsMember(I) Then
                        picSelected.Line (rcDraw.Left - 2, rcDraw.Top - 2)-(rcDraw.Right + 2, rcDraw.Bottom + 2), vbBlue, BF
                        picSelected.PaintPicture ExtractLocalTile(I, True), rcDraw.Left, rcDraw.Top
                    Else
                        picSelected.PaintPicture ExtractLocalTile(I), rcDraw.Left, rcDraw.Top
                    End If
                End If
                DrawIndex = DrawIndex + 1
            End With
        End If
    Next
    
    VScrollSelected.Max = YMax
    
End Sub

Public Sub PaintTileset()
    Dim rcDraw As RECT
    Dim I As Integer
    Dim YMax As Long
    Dim TileCount As Integer
    
    On Error GoTo PaintErr
    
    picTileset.Cls
    
    With EditTileMatch.TSDef
        If Not .IsLoaded Then
            .Load
        End If
        TileCount = (ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth) * (ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight)
    End With
    
    For I = 0 To TileCount - 1
        With EditTileMatch.TSDef
            rcDraw = GetTileRect(I)
            If rcDraw.Top > YMax Then YMax = rcDraw.Top
            rcDraw.Top = rcDraw.Top - vscrollTileset.Value
            rcDraw.Bottom = rcDraw.Bottom - vscrollTileset.Value
            If rcDraw.Bottom > 0 And rcDraw.Top < picTileset.ScaleHeight Then
                If Highlighted.IsMember(I) Then
                    picTileset.Line (rcDraw.Left - 2, rcDraw.Top - 2)-(rcDraw.Right + 2, rcDraw.Bottom + 2), vbBlue, BF
                    picTileset.PaintPicture ExtractLocalTile(I, True), rcDraw.Left, rcDraw.Top
                Else
                    picTileset.PaintPicture ExtractLocalTile(I), rcDraw.Left, rcDraw.Top
                End If
            End If
        End With
    Next
    
    vscrollTileset.Max = YMax
    Exit Sub
    
PaintErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function ExtractLocalTile(ByVal Index As Integer, Optional bHighlight As Boolean = False) As StdPicture
    Dim TileCols As Integer

    With EditTileMatch.TSDef
        TileCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        If Index < 0 Then Exit Function
        Set ExtractLocalTile = ExtractTile(.Image, .TileWidth * (Index Mod TileCols), .TileHeight * (Index \ TileCols), .TileWidth, .TileHeight, bHighlight)
    End With
    
End Function

Private Function GetTileRect(Index) As RECT
    Dim FitCols As Integer
        
    With EditTileMatch.TSDef
        FitCols = (picGroup.ScaleWidth) \ (.TileWidth + 6)
        GetTileRect.Left = (Index Mod FitCols) * (.TileWidth + 6) + 3
        GetTileRect.Top = (Index \ FitCols) * (.TileHeight + 6) + 3
        GetTileRect.Right = GetTileRect.Left + .TileWidth - 1
        GetTileRect.Bottom = GetTileRect.Top + .TileHeight - 1
    End With
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "TileMatching", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub imgMatchTile_Click(Index As Integer)
    If SelGroup = Index Then
        SelGroup = -1
    Else
        SelGroup = Index
    End If
    
    Highlighted.ClearAll
    UpdateImages
    Me.Refresh
    PaintSelected
    
End Sub

Private Sub UpdateImages()
    Dim I As Integer
    
    For I = imgMatchTile.LBound To imgMatchTile.UBound
        If EditTileMatch.TileMatches.MatchGroup(I).IsEmpty Then
            Set imgMatchTile(I).Picture = Nothing
        Else
            If I = SelGroup Then
                Set imgMatchTile(I).Picture = ExtractLocalTile(EditTileMatch.TileMatches.MatchGroup(I).GetMember(0), True)
            Else
                Set imgMatchTile(I).Picture = ExtractLocalTile(EditTileMatch.TileMatches.MatchGroup(I).GetMember(0), False)
            End If
        End If
    Next I
    
End Sub

Private Sub imgMatchTile_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            EditTileMatch.TileMatches.MatchGroup(Index).SetMember CInt(DropTiles(I))
        Next
        EditTileMatch.UpdateTotalGroup
    End If

    Prj.IsDirty = True
    UpdateImages
    Me.Refresh
    PaintSelected
    PaintGroup
    PaintTileset

End Sub

Private Sub imgMatchTile_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectMove Or vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub picGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 0
End Sub

Private Sub HandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Grp As Integer)
    Dim Idx As Integer

    If Grp = 0 Then
        Idx = GetXYGrpIndex(X, Y)
    ElseIf Grp = 1 Then
        Idx = GetXYSelIndex(X, Y)
    ElseIf Grp = 2 Then
        Idx = GetXYTilesetIndex(X, Y)
    End If
    
    If Idx >= 0 Then
        If Shift And vbCtrlMask Then
            If Highlighted.IsMember(Idx) Then
                Highlighted.ClearMember Idx
            Else
                Highlighted.SetMember Idx
                Set DragPic = ExtractLocalTile(Idx)
                DragState = 1
                StartDragPt.X = X
                StartDragPt.Y = Y
            End If
            PaintSelected
            PaintGroup
            PaintTileset
        Else
            If Not (Highlighted.IsMember(Idx)) Then
                SelectSingle Idx
            End If
            Set DragPic = ExtractLocalTile(Idx)
            DragState = 1
            StartDragPt.X = X
            StartDragPt.Y = Y
        End If
    Else
        If (Shift And vbCtrlMask) = 0 Then
            If Not Highlighted.IsEmpty Then
                Highlighted.ClearAll
                PaintSelected
                PaintGroup
                PaintTileset
            End If
        End If
        DragState = 0
    End If
    
End Sub

Private Function GetXYSelIndex(X As Single, Y As Single) As Integer
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim V As Variant

    V = EditTileMatch.AllTiles.GetArray
    If IsEmpty(V) Then
        GetXYSelIndex = -1
        Exit Function
    End If
    For J = LBound(V) To UBound(V)
        I = V(J)
        If EditTileMatch.TileMatches.MatchGroup(SelGroup).IsMember(I) Then
            With EditTileMatch.TSDef
                rcDraw = GetTileRect(DrawIndex)
                rcDraw.Top = rcDraw.Top - VScrollSelected.Value
                rcDraw.Bottom = rcDraw.Bottom - VScrollSelected.Value
                If X >= rcDraw.Left And X <= rcDraw.Right And Y <= rcDraw.Bottom And Y >= rcDraw.Top Then
                    GetXYSelIndex = I
                    Exit Function
                End If
                DrawIndex = DrawIndex + 1
            End With
        End If
    Next
        
    GetXYSelIndex = -1
    
End Function

Private Function GetXYGrpIndex(X As Single, Y As Single) As Integer
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim V As Variant
    
    V = EditTileMatch.AllTiles.GetArray
    If IsEmpty(V) Then
        GetXYGrpIndex = -1
        Exit Function
    End If
    For J = LBound(V) To UBound(V)
        I = V(J)
        If Not EditTileMatch.TileMatches.IsMember(I) Then
            With EditTileMatch.TSDef
                rcDraw = GetTileRect(DrawIndex)
                rcDraw.Top = rcDraw.Top - VScrollGroup.Value
                rcDraw.Bottom = rcDraw.Bottom - VScrollGroup.Value
                If X >= rcDraw.Left And X <= rcDraw.Right And Y >= rcDraw.Top And Y <= rcDraw.Bottom Then
                    GetXYGrpIndex = I
                    Exit Function
                End If
                DrawIndex = DrawIndex + 1
            End With
        End If
    Next
    
    GetXYGrpIndex = -1
    
End Function

Private Function GetXYTilesetIndex(X As Single, Y As Single) As Integer
    Dim rcDraw As RECT
    Dim I As Integer
    Dim YMax As Long
    Dim TileCount As Integer
    
    With EditTileMatch.TSDef
        TileCount = (ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth) * (ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight)
    End With
    
    For I = 0 To TileCount - 1
        With EditTileMatch.TSDef
            rcDraw = GetTileRect(I)
            rcDraw.Top = rcDraw.Top - vscrollTileset.Value
            rcDraw.Bottom = rcDraw.Bottom - vscrollTileset.Value
            If X >= rcDraw.Left And X <= rcDraw.Right And Y >= rcDraw.Top And Y <= rcDraw.Bottom Then
                GetXYTilesetIndex = I
                Exit Function
            End If
        End With
    Next
    
    GetXYTilesetIndex = -1

End Function

Private Sub picGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragState = 1 Then
        If Abs(StartDragPt.X - X) > 3 Or Abs(StartDragPt.Y - Y) > 3 Then
            picGroup.OLEDrag
            DragState = 2
        End If
    End If
End Sub

Private Sub picGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Idx As Integer
    
    Idx = GetXYGrpIndex(X, Y)
    
    If DragState = 1 And Idx >= 0 Then
        If (Shift And vbCtrlMask) = 0 Then
            SelectSingle Idx
        End If
    End If
    
    DragState = 0
    
End Sub

Private Sub picGroup_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            EditTileMatch.AllTiles.SetMember CInt(DropTiles(I))
        Next
    End If

    Prj.IsDirty = True
    UpdateImages
    Me.Refresh
    PaintSelected
    PaintGroup
    PaintTileset
    
End Sub

Private Sub picGroup_OLESetData(Data As DataObject, DataFormat As Integer)
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

Private Sub picGroup_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Data.SetData Format:=vbCFDIB
    Data.SetData Format:=CInt("&H" & Hex$(CFTiles))
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
End Sub

Private Sub picGroup_Paint()
    PaintGroup
End Sub

Private Sub picSelected_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 1
End Sub

Private Sub picSelected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragState = 1 Then
        If Abs(StartDragPt.X - X) > 3 Or Abs(StartDragPt.Y - Y) > 3 Then
            picSelected.OLEDrag
            DragState = 2
        End If
    End If
End Sub

Private Sub picSelected_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Idx As Integer
    
    Idx = GetXYSelIndex(X, Y)
    
    If DragState = 1 And Idx >= 0 Then
        If (Shift And vbCtrlMask) = 0 Then
            SelectSingle Idx
        End If
    End If
    
    DragState = 0
    
End Sub

Private Sub picSelected_OLECompleteDrag(Effect As Long)
    Dim I As Integer
    Dim V As Variant
    
    If (Effect And vbDropEffectMove) = vbDropEffectMove Then
        V = Highlighted.GetArray
        If IsEmpty(V) Then Exit Sub
        For I = LBound(V) To UBound(V)
            EditTileMatch.TileMatches.MatchGroup(SelGroup).ClearMember CInt(V(I))
        Next
    End If
    
    UpdateImages
    Me.Refresh
    PaintSelected
    PaintGroup
    
End Sub

Private Sub picSelected_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            EditTileMatch.TileMatches.MatchGroup(SelGroup).SetMember CInt(DropTiles(I))
        Next
        EditTileMatch.UpdateTotalGroup
    End If

    Prj.IsDirty = True
    UpdateImages
    Me.Refresh
    PaintSelected
    PaintGroup
    PaintTileset
    
End Sub

Private Sub picSelected_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectMove Or vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub picSelected_OLESetData(Data As DataObject, DataFormat As Integer)
    picGroup_OLESetData Data, DataFormat
End Sub

Private Sub picSelected_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    picGroup_OLEStartDrag Data, AllowedEffects
    AllowedEffects = vbDropEffectMove
End Sub

Private Sub picSelected_Paint()
    PaintSelected
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 2
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
    
    Idx = GetXYTilesetIndex(X, Y)
    
    If DragState = 1 And Idx >= 0 Then
        If (Shift And vbCtrlMask) = 0 Then
            SelectSingle Idx
        End If
    End If
    
    DragState = 0
    
End Sub

Private Sub SelectSingle(Idx As Integer)
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
        PaintSelected
        PaintGroup
    End If
    
End Sub

Private Sub picTileset_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    Dim J As Integer
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            EditTileMatch.AllTiles.ClearMember CInt(DropTiles(I))
            For J = 0 To 14
                EditTileMatch.TileMatches.MatchGroup(J).ClearMember CInt(DropTiles(I))
            Next
        Next
    End If

    Prj.IsDirty = True
    UpdateImages
    Me.Refresh
    PaintSelected
    PaintGroup
    PaintTileset

End Sub

Private Sub picTileset_OLESetData(Data As DataObject, DataFormat As Integer)
    picGroup_OLESetData Data, DataFormat
End Sub

Private Sub picTileset_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    picGroup_OLEStartDrag Data, AllowedEffects
    AllowedEffects = vbDropEffectCopy
End Sub

Private Sub picTileset_Paint()
    PaintTileset
End Sub

Private Sub VScrollGroup_Change()
    picGroup.Refresh
End Sub

Private Sub VScrollGroup_Scroll()
    picGroup.Refresh
End Sub

Private Sub VScrollSelected_Change()
    picSelected.Refresh
End Sub

Private Sub VScrollSelected_Scroll()
    picSelected.Refresh
End Sub

Private Sub vscrollTileset_Change()
    picTileset.Refresh
End Sub

Private Sub vscrollTileset_Scroll()
    picTileset.Refresh
End Sub
