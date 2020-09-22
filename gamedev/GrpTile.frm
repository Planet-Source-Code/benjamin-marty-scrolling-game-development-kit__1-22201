VERSION 5.00
Begin VB.Form frmGroupTiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Categories"
   ClientHeight    =   6345
   ClientLeft      =   1125
   ClientTop       =   390
   ClientWidth     =   5295
   HelpContextID   =   103
   Icon            =   "GrpTile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSolidity 
      Caption         =   "Create/Edit Solidity Definitions for Selected Tileset"
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   5055
      Begin VB.CommandButton cmdSaveSolidity 
         Caption         =   "Save"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cboDownCeil 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cboUpCeil 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cboDownhill 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboUphill 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboSolid 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeleteSolidity 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cboSolidityName 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblDownCeil 
         Caption         =   "Ceiling down tile category:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblUpCeil 
         Caption         =   "Ceiling up tile category:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblDownhill 
         Caption         =   "Downhill tile category:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblSolid 
         Caption         =   "Solid tile category:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblUphill 
         Caption         =   "Uphill tile category:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblSolidName 
         Caption         =   "Solidity Definition Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Category"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cboTileset 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox cboCurGroup 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.VScrollBar VScrollTileSet 
      Height          =   975
      LargeChange     =   30
      Left            =   4920
      SmallChange     =   5
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picTileSet 
      Height          =   975
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   10
      Top             =   2520
      Width           =   4815
   End
   Begin VB.PictureBox picGroup 
      Height          =   975
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   7
      Top             =   1200
      Width           =   4815
   End
   Begin VB.VScrollBar VScrollGroup 
      Height          =   975
      LargeChange     =   30
      Left            =   4920
      SmallChange     =   5
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblSelTileset 
      BackStyle       =   0  'Transparent
      Caption         =   "Tileset:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblCurGrp 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Category:"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles available in tileset:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles in this category:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "frmGroupTiles"
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
' File: GrpTile.frm - Tile Category and Solidity Management Dialog
'
'======================================================================

Option Explicit

Public Highlighted As New TileGroup
Dim SelGroup As Integer
Dim DragState As Integer
Dim DragPic As StdPicture
Dim CFTiles As Long
Dim StartDragPt As POINTAPI

Private Sub cboCurGroup_Click()
    PaintGroup
End Sub

Private Sub cboSolidityName_Change()
    If Prj.SolidDefExists(cboSolidityName.Text, cboTileset.Text) Then LoadSolidityDef Prj.SolidDefs(cboSolidityName.Text, cboTileset.Text)
End Sub

Sub LoadSolidityDef(SD As SolidDef)
    SelectGroupInCombo cboSolid, SD.Solid
    SelectGroupInCombo cboUphill, SD.Uphill
    SelectGroupInCombo cboDownhill, SD.Downhill
    SelectGroupInCombo cboUpCeil, SD.UpCeil
    SelectGroupInCombo cboDownCeil, SD.DownCeil
End Sub

Sub SelectGroupInCombo(C As ComboBox, Grp As Category)
    If Grp Is Nothing Then
        C.ListIndex = 0
    Else
        C.ListIndex = Grp.GetIndexByTileset(Grp.TSName) + 1
    End If
End Sub

Private Sub cboSolidityName_Click()
    cboSolidityName_Change
End Sub

Private Sub cboTileset_Click()
    LoadGroups
    LoadSolidDefs
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cboCurGroup.ListIndex < 0 Then
        MsgBox "Please select an existing category before selecting this command."
    Else
        Prj.RemoveGroup cboCurGroup.Text, cboTileset.Text
        cboCurGroup.Text = ""
    End If
    LoadGroups
    LoadSolidDefs
End Sub

Private Sub cmdDeleteSolidity_Click()
    If Prj.SolidDefExists(cboSolidityName.Text, cboTileset.Text) Then
        Prj.RemoveSolidDef Prj.SolidDefs(cboSolidityName.Text, cboTileset.Text)
    End If
    LoadSolidDefs
End Sub

Private Sub cmdSaveSolidity_Click()
    Dim SD As SolidDef

    If Prj.SolidDefExists(cboSolidityName.Text, cboTileset.Text) Then
        Set SD = Prj.SolidDefs(cboSolidityName.Text, cboTileset.Text)
    Else
        Set SD = Prj.AddSolidDef(cboSolidityName.Text, cboTileset.Text)
    End If
    
    If Prj.GroupExists(cboSolid.Text, SD.TSName) Then
        Set SD.Solid = Prj.Groups(cboSolid.Text, SD.TSName)
    Else
        Set SD.Solid = Nothing
    End If
    If Prj.GroupExists(cboUphill.Text, SD.TSName) Then
        Set SD.Uphill = Prj.Groups(cboUphill.Text, SD.TSName)
    Else
        Set SD.Uphill = Nothing
    End If
    If Prj.GroupExists(cboDownhill.Text, SD.TSName) Then
        Set SD.Downhill = Prj.Groups(cboDownhill.Text, SD.TSName)
    Else
        Set SD.Downhill = Nothing
    End If
    If Prj.GroupExists(cboUpCeil.Text, SD.TSName) Then
        Set SD.UpCeil = Prj.Groups(cboUpCeil.Text, SD.TSName)
    Else
        Set SD.UpCeil = Nothing
    End If
    If Prj.GroupExists(cboDownCeil.Text, SD.TSName) Then
        Set SD.DownCeil = Prj.Groups(cboDownCeil.Text, SD.TSName)
    Else
        Set SD.DownCeil = Nothing
    End If

    Prj.IsDirty = True
    LoadSolidDefs
End Sub

Private Sub Form_Initialize()
    CFTiles = RegisterClipboardFormat("TileGroup")
End Sub

Public Function MakeGroupExist() As Category
    If Len(cboCurGroup.Text) = 0 Then Exit Function
    If Not (Prj.GroupExists(cboCurGroup.Text, cboTileset.Text)) Then
        Set MakeGroupExist = Prj.AddGroup(cboCurGroup.Text, cboTileset.Text)
        LoadGroups
        LoadSolidDefs
    Else
        Set MakeGroupExist = Prj.Groups(cboCurGroup.Text, cboTileset.Text)
    End If
End Function

Public Sub PaintGroup()
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim YMax As Long
    Dim V As Variant
    Dim Grp As Category
    
    picGroup.Cls
    
    If cboTileset.ListIndex < 0 Or Len(cboCurGroup.Text) <= 0 Then Exit Sub
    
    Set Grp = MakeGroupExist
    V = Grp.Group.GetArray
    If IsEmpty(V) Then Exit Sub
    For J = LBound(V) To UBound(V)
        I = V(J)
        rcDraw = GetTileRect(DrawIndex)
        If rcDraw.Top > YMax Then YMax = rcDraw.Top
        rcDraw.Top = rcDraw.Top - VScrollGroup.Value
        rcDraw.Bottom = rcDraw.Bottom - VScrollGroup.Value
        If rcDraw.Bottom > 0 And rcDraw.Top < picGroup.ScaleHeight Then
            If Highlighted.IsMember(I) Then
                picGroup.Line (rcDraw.Left - 2, rcDraw.Top - 2)-(rcDraw.Right + 2, rcDraw.Bottom + 2), vbBlue, BF
                picGroup.PaintPicture ExtractLocalTile(I, True), rcDraw.Left, rcDraw.Top
            Else
                picGroup.PaintPicture ExtractLocalTile(I), rcDraw.Left, rcDraw.Top
            End If
        End If
        DrawIndex = DrawIndex + 1
    Next
    
    VScrollGroup.Max = YMax
    
End Sub

Public Sub PaintTileset()
    Dim rcDraw As RECT
    Dim I As Integer
    Dim YMax As Long
    Dim TileCount As Integer
    
    picTileset.Cls
    
    If cboTileset.ListIndex < 0 Then Exit Sub
    
    With Prj.TileSetDef(cboTileset.Text)
        If .Image Is Nothing Then
            .Load
        End If
        TileCount = (ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth) * (ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight)
    End With
    
    For I = 0 To TileCount - 1
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
    Next
    
    vscrollTileset.Max = YMax

End Sub

Private Function ExtractLocalTile(ByVal Index As Integer, Optional bHighlight As Boolean = False) As StdPicture
    Dim TileCols As Integer

    If cboTileset.ListIndex < 0 Then Exit Function

    With Prj.TileSetDef(cboTileset.Text)
        TileCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        If Index < 0 Then Exit Function
        Set ExtractLocalTile = ExtractTile(.Image, .TileWidth * (Index Mod TileCols), .TileHeight * (Index \ TileCols), .TileWidth, .TileHeight, bHighlight)
    End With
    
End Function

Private Function GetTileRect(Index) As RECT
    Dim FitCols As Integer
    
    With Prj.TileSetDef(cboTileset.Text)
        FitCols = (picGroup.ScaleWidth) \ (.TileWidth + 6)
        GetTileRect.Left = (Index Mod FitCols) * (.TileWidth + 6) + 3
        GetTileRect.Top = (Index \ FitCols) * (.TileHeight + 6) + 3
        GetTileRect.Right = GetTileRect.Left + .TileWidth - 1
        GetTileRect.Bottom = GetTileRect.Top + .TileHeight - 1
    End With
    
End Function

Private Sub Form_Load()
    Dim I As Integer
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "TileCategories", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    cboTileset.Clear
    For I = 0 To Prj.TileSetDefCount - 1
        cboTileset.AddItem Prj.TileSetDef(I).Name
    Next I
    Exit Sub

LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub LoadGroups()
    Dim I As Integer
    Dim PrevText As String
    
    PrevText = cboCurGroup.Text
    cboCurGroup.Clear
    For I = 0 To Prj.GroupByTilesetCount(cboTileset.Text) - 1
        cboCurGroup.AddItem Prj.TilesetGroupByIndex(cboTileset.Text, I).Name
    Next I
    cboCurGroup.Text = PrevText
    PaintTileset
    PaintGroup
End Sub

Sub LoadSolidDefs()
    Dim I As Integer
    
    cboSolidityName.Clear
    For I = 0 To Prj.SolidDefByTilesetCount(cboTileset.Text) - 1
        cboSolidityName.AddItem Prj.SolidDefsByIndex(Prj.SolidDefIndexByTileset(cboTileset.Text, I)).Name
    Next
    
    LoadComboWithRelaventGroups cboSolid
    LoadComboWithRelaventGroups cboUphill
    LoadComboWithRelaventGroups cboDownhill
    LoadComboWithRelaventGroups cboUpCeil
    LoadComboWithRelaventGroups cboDownCeil
End Sub

Sub LoadComboWithRelaventGroups(C As ComboBox)
    Dim I As Integer
    
    C.Clear
    
    C.AddItem "<none>"
    For I = 0 To Prj.GroupByTilesetCount(cboTileset.Text) - 1
        C.AddItem Prj.TilesetGroupByIndex(cboTileset.Text, I).Name
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "TileCategories", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub picGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y, 0
End Sub

Private Sub HandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Grp As Integer)
    Dim Idx As Integer

    If Grp = 0 Then
        Idx = GetXYGrpIndex(X, Y)
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
                PaintGroup
                PaintTileset
            End If
        End If
        DragState = 0
    End If
    
End Sub

Private Function GetXYGrpIndex(X As Single, Y As Single) As Integer
    Dim DrawIndex As Integer
    Dim rcDraw As RECT
    Dim I As Integer, J As Integer
    Dim V As Variant
    Dim Grp As Category
    
    Set Grp = MakeGroupExist
    If Grp Is Nothing Then
        GetXYGrpIndex = -1
        Exit Function
    End If
    V = Grp.Group.GetArray
    If IsEmpty(V) Then
        GetXYGrpIndex = -1
        Exit Function
    End If
    For J = LBound(V) To UBound(V)
        I = V(J)
        rcDraw = GetTileRect(DrawIndex)
        rcDraw.Top = rcDraw.Top - VScrollGroup.Value
        rcDraw.Bottom = rcDraw.Bottom - VScrollGroup.Value
        If X >= rcDraw.Left And X <= rcDraw.Right And Y >= rcDraw.Top And Y <= rcDraw.Bottom Then
            GetXYGrpIndex = I
            Exit Function
        End If
        DrawIndex = DrawIndex + 1
    Next
    
    GetXYGrpIndex = -1
    
End Function

Private Function GetXYTilesetIndex(X As Single, Y As Single) As Integer
    Dim rcDraw As RECT
    Dim I As Integer
    Dim YMax As Long
    Dim TileCount As Integer
    
    If cboTileset.ListIndex < 0 Then
        GetXYTilesetIndex = -1
        Exit Function
    End If
    With Prj.TileSetDef(cboTileset.Text)
        TileCount = (ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth) * (ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight)
    End With
    
    For I = 0 To TileCount - 1
        rcDraw = GetTileRect(I)
        rcDraw.Top = rcDraw.Top - vscrollTileset.Value
        rcDraw.Bottom = rcDraw.Bottom - vscrollTileset.Value
        If X >= rcDraw.Left And X <= rcDraw.Right And Y >= rcDraw.Top And Y <= rcDraw.Bottom Then
            GetXYTilesetIndex = I
            Exit Function
        End If
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
    Dim Grp As Category
    
    If cboTileset.ListIndex < 0 Or Len(cboCurGroup.Text) <= 0 Then Exit Sub
    
    Set Grp = MakeGroupExist
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            Grp.Group.SetMember CInt(DropTiles(I))
        Next
    End If

    Prj.IsDirty = True
    Me.Refresh
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
        PaintGroup
    End If
    
End Sub

Private Sub picTileset_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DropTiles As Variant
    Dim TileCount As Byte
    Dim I As Integer
    Dim J As Integer
    Dim Grp As Category
    
    If cboTileset.ListIndex < 0 Or Len(cboCurGroup.Text) <= 0 Then Exit Sub
    
    Set Grp = MakeGroupExist
    
    If Data.GetFormat(CInt("&h" & Hex$(CFTiles))) Then
        DropTiles = Data.GetData(CInt("&H" & Hex$(CFTiles)))
        TileCount = DropTiles(LBound(DropTiles))
        For I = LBound(DropTiles) + 1 To LBound(DropTiles) + TileCount
            Grp.Group.ClearMember CInt(DropTiles(I))
        Next
    End If

    Prj.IsDirty = True
    Me.Refresh
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

Private Sub vscrollTileset_Change()
    picTileset.Refresh
End Sub

Private Sub vscrollTileset_Scroll()
    picTileset.Refresh
End Sub
