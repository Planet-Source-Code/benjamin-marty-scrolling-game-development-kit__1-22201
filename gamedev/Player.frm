VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Settings"
   ClientHeight    =   5760
   ClientLeft      =   2505
   ClientTop       =   330
   ClientWidth     =   6495
   HelpContextID   =   111
   Icon            =   "Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraInventory 
      Caption         =   "Inventory"
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   6255
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   2490
      End
      Begin VB.TextBox txtTileIndex 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   945
      End
      Begin VB.ComboBox cboTileset 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtMaxQuantity 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtInitQuantity 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cboQuantityDisplay 
         Height          =   315
         ItemData        =   "Player.frx":0442
         Left            =   1440
         List            =   "Player.frx":047C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtIconCountPerRepeat 
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdBrowseBarColor 
         Caption         =   "..."
         Height          =   255
         Left            =   1470
         TabIndex        =   26
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdBrowseBarBackground 
         Caption         =   "..."
         Height          =   255
         Left            =   3510
         TabIndex        =   28
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdBrowseBarOutline 
         Caption         =   "..."
         Height          =   255
         Left            =   5190
         TabIndex        =   30
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtBarLength 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox txtBarThickness 
         Height          =   285
         Left            =   3720
         TabIndex        =   36
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox txtItemX 
         Height          =   285
         Left            =   3120
         TabIndex        =   39
         Top             =   2880
         Width           =   585
      End
      Begin VB.TextBox txtItemY 
         Height          =   285
         Left            =   4800
         TabIndex        =   42
         Top             =   2880
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   45
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   46
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   47
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4560
         TabIndex        =   49
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Update"
         Height          =   375
         Left            =   5400
         TabIndex        =   50
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox chkNoOutline 
         Caption         =   "None"
         Height          =   255
         Left            =   5445
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   615
      End
      Begin MSComCtl2.UpDown updItemY 
         Height          =   285
         Left            =   5400
         TabIndex        =   43
         Top             =   2880
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtItemY"
         BuddyDispid     =   196623
         OrigLeft        =   5400
         OrigTop         =   2880
         OrigRight       =   5595
         OrigBottom      =   3135
         Increment       =   5
         Max             =   479
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updItemX 
         Height          =   270
         Left            =   3720
         TabIndex        =   40
         Top             =   2880
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtItemX"
         BuddyDispid     =   196622
         OrigLeft        =   3720
         OrigTop         =   2880
         OrigRight       =   3915
         OrigBottom      =   3135
         Increment       =   5
         Max             =   639
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updBarThickness 
         Height          =   285
         Left            =   4261
         TabIndex        =   37
         Top             =   2520
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBarThickness"
         BuddyDispid     =   196621
         OrigLeft        =   4200
         OrigTop         =   2520
         OrigRight       =   4395
         OrigBottom      =   2775
         Max             =   32
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updBarLength 
         Height          =   285
         Left            =   1861
         TabIndex        =   34
         Top             =   2520
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         Value           =   3
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBarLength"
         BuddyDispid     =   196620
         OrigLeft        =   1800
         OrigTop         =   2520
         OrigRight       =   1995
         OrigBottom      =   2775
         Max             =   600
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updIconIndex 
         Height          =   285
         Left            =   5880
         TabIndex        =   16
         Top             =   720
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTileIndex"
         BuddyDispid     =   196611
         OrigLeft        =   2325
         OrigTop         =   1035
         OrigRight       =   2520
         OrigBottom      =   1410
         Max             =   255
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblIconTileset 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon Tileset:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image imgTilePreview 
         Height          =   975
         Left            =   4920
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIconIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon Index:"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblMaxOwn 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Quantity:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblInitOwn 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Quantity:"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblQuantityDisplay 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity Display:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblM 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon count per repetiton (M):"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Shape shpBarColor 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1200
         Top             =   2160
         Width           =   255
      End
      Begin VB.Shape shpBarBackground 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3240
         Top             =   2160
         Width           =   255
      End
      Begin VB.Shape shpBarOutline 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4920
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblBarLength 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Length:"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblBarThickness 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Thickness:"
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblItemPosition 
         BackStyle       =   0  'Transparent
         Caption         =   "Coordinates to display inventory item:  X:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label lblItemY 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblBarOutline 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Outline:"
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblBarBackground 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Background:"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label lblBarColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Color:"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2160
         Width           =   885
      End
   End
   Begin VB.TextBox txtQuantityMargin 
      Height          =   285
      Left            =   5040
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   51
      Top             =   5280
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtScrollY 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtScrollX 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox cboMaps 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblBarMargin 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Margin (between icon and bars):"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblScrollY 
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical:"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblScrollX 
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblScrollMargin 
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll margins (distance between player and edge):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblStartMap 
      BackStyle       =   0  'Transparent
      Caption         =   "Start on Map:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlayer"
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
' File: Player.frm - Game Player and Inventory Dialog
'
'======================================================================

Option Explicit

Public nCurInv As Integer

Sub FillMaps()
    Dim Idx As Integer
    
    cboMaps.Clear
    For Idx = 0 To Prj.MapCount - 1
        cboMaps.AddItem Prj.Maps(Idx).Name
    Next Idx
End Sub

Sub FillTilesets()
    Dim Idx As Integer
    
    cboTileset.Clear
    cboTileset.AddItem "(none)"
    For Idx = 0 To Prj.TileSetDefCount - 1
        cboTileset.AddItem Prj.TileSetDef(Idx).Name
    Next
End Sub

Private Sub cboQuantityDisplay_Change()
    Dim bEnableRepeat As Boolean
    Dim bEnableBars As Boolean
    
    Select Case cboQuantityDisplay.ListIndex
    Case QuantityDisplay.QD_ICONINDEXREPEATABOVE, QuantityDisplay.QD_ICONINDEXREPEATBELOW, QuantityDisplay.QD_ICONINDEXREPEATRIGHT
        bEnableRepeat = True
        bEnableBars = False
    Case QuantityDisplay.QD_HORZBARABOVE, QuantityDisplay.QD_HORZBARBELOW, QuantityDisplay.QD_HORZBARRIGHT, _
         QuantityDisplay.QD_VERTBARABOVE, QuantityDisplay.QD_VERTBARLEFT, QuantityDisplay.QD_VERTBARRIGHT
        bEnableRepeat = False
        bEnableBars = True
    Case Else
        bEnableRepeat = False
        bEnableBars = False
    End Select
    lblM.Enabled = bEnableRepeat
    txtIconCountPerRepeat.Enabled = bEnableRepeat
    txtIconCountPerRepeat.BackColor = IIf(bEnableRepeat, &H80000005, &H8000000F)
    lblBarColor.Enabled = bEnableBars
    cmdBrowseBarColor.Enabled = bEnableBars
    lblBarBackground.Enabled = bEnableBars
    cmdBrowseBarBackground.Enabled = bEnableBars
    lblBarOutline.Enabled = bEnableBars
    cmdBrowseBarOutline.Enabled = bEnableBars
    chkNoOutline.Enabled = bEnableBars
    lblBarLength.Enabled = bEnableBars
    txtBarLength.Enabled = bEnableBars
    txtBarLength.BackColor = IIf(bEnableBars, &H80000005, &H8000000F)
    updBarLength.Enabled = bEnableBars
    lblBarThickness.Enabled = bEnableBars
    txtBarThickness.Enabled = bEnableBars
    txtBarThickness.BackColor = IIf(bEnableBars, &H80000005, &H8000000F)
    updBarThickness.Enabled = bEnableBars
    
End Sub

Private Sub cboQuantityDisplay_Click()
    cboQuantityDisplay_Change
End Sub

Private Sub cboTileset_Change()
    On Error Resume Next
    If cboTileset.ListIndex <= 0 Then Exit Sub
    With Prj.TileSetDef(cboTileset.List(cboTileset.ListIndex))
        updIconIndex.Max = (ScaleX(.Image.Width, vbHimetric, vbPixels) / .TileWidth) * _
                           (ScaleY(.Image.Height, vbHimetric, vbPixels) / .TileHeight) - 1
    End With
    If Err.Number Then updIconIndex.Max = 255
End Sub

Private Sub cboTileset_Click()
    cboTileset_Change
End Sub

Private Sub chkNoOutline_Click()
    If chkNoOutline.Value = vbChecked Then
        shpBarOutline.Visible = False
        cmdBrowseBarOutline.Enabled = False
    Else
        shpBarOutline.Visible = True
        cmdBrowseBarOutline.Enabled = True
    End If
End Sub

Private Sub cmdBrowseBarBackground_Click()
    dlgColor.Color = shpBarBackground.FillColor
    dlgColor.Flags = cdlCCRGBInit
    dlgColor.ShowColor
    shpBarBackground.FillColor = dlgColor.Color
End Sub

Private Sub cmdBrowseBarColor_Click()
    dlgColor.Color = shpBarColor.FillColor
    dlgColor.Flags = cdlCCRGBInit
    dlgColor.ShowColor
    shpBarColor.FillColor = dlgColor.Color
End Sub

Private Sub cmdBrowseBarOutline_Click()
    dlgColor.Color = shpBarOutline.FillColor
    dlgColor.Flags = cdlCCRGBInit
    dlgColor.ShowColor
    shpBarOutline.FillColor = dlgColor.Color
End Sub

Private Sub cmdDelete_Click()
    Prj.GamePlayer.RemoveInventoryItem nCurInv
    If nCurInv > Prj.GamePlayer.InventoryCount - 1 Then
        nCurInv = Prj.GamePlayer.InventoryCount - 1
    End If
    LoadInventory
    Prj.IsDirty = True
End Sub

Private Sub cmdFirst_Click()
    StoreInventory
    nCurInv = 0
    LoadInventory
End Sub

Private Sub cmdLast_Click()
    StoreInventory
    nCurInv = Prj.GamePlayer.InventoryCount - 1
    LoadInventory
End Sub

Private Sub cmdNew_Click()
    nCurInv = Prj.GamePlayer.InventoryCount
    If cboQuantityDisplay.ListIndex < 0 Then cboQuantityDisplay.ListIndex = 0
    If Not IsNumeric(txtItemX.Text) Then
        txtItemX.Text = "0"
    End If
    If Not IsNumeric(txtItemY.Text) Then
        txtItemY.Text = "0"
    End If
    Prj.GamePlayer.AddInventoryItem cboQuantityDisplay.ListIndex, CInt(txtItemX.Text), CInt(txtItemY.Text)
    StoreInventory
    LoadInventory
End Sub

Private Sub cmdNext_Click()
    StoreInventory
    nCurInv = nCurInv + 1
    LoadInventory
End Sub

Private Sub cmdOK_Click()
    StorePlayer
    Unload Me
End Sub

Private Sub cmdPrevious_Click()
    StoreInventory
    nCurInv = nCurInv - 1
    LoadInventory
End Sub

Private Sub cmdSave_Click()
    StoreInventory
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "PlayerSettings", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    FillMaps
    FillTilesets
    Prj.GamePlayer.ReIndexTilesetRefs
    LoadPlayer
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub LoadPlayer()
    Dim Idx As Integer
    Dim SprIdx As Integer
    
    On Error GoTo LoadPlayerErr
    
    With Prj.GamePlayer
        For Idx = 0 To cboMaps.ListCount - 1
            If cboMaps.List(Idx) = .StartMapName Then
                cboMaps.ListIndex = Idx
                Exit For
            End If
        Next
        txtScrollX.Text = CStr(.ScrollMarginX)
        txtScrollY.Text = CStr(.ScrollMarginY)
        txtQuantityMargin.Text = CStr(.InvBarMargin)
        LoadInventory
    End With
    Exit Sub
    
LoadPlayerErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub StorePlayer()
    On Error GoTo StorePlayerErr
    
    Prj.IsDirty = True
    With Prj.GamePlayer
        .StartMapName = cboMaps.List(cboMaps.ListIndex)
        .ScrollMarginX = CInt(txtScrollX.Text)
        .ScrollMarginY = CInt(txtScrollY.Text)
        .InvBarMargin = CInt(txtQuantityMargin.Text)
        StoreInventory
    End With
    Exit Sub
    
StorePlayerErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub LoadInventory()
    Dim Idx As Integer
    
    On Error GoTo LoadInvErr
    With Prj.GamePlayer
        If .InventoryCount <= 0 Then
            fraInventory.Caption = "No Inventory Defined"
            cmdFirst.Enabled = False
            cmdPrevious.Enabled = False
            cmdNext.Enabled = False
            cmdLast.Enabled = False
            cmdDelete.Enabled = False
            cmdSave.Enabled = False
            Exit Sub
        End If
        fraInventory.Caption = "Inventory Item " & CStr(nCurInv + 1) & " of " & .InventoryCount
        txtItemName.Text = .InventoryItemName(nCurInv)
        If .InvIconTilesetIdx(nCurInv) < 0 Then
            cboTileset.ListIndex = 0
        Else
            For Idx = 0 To cboTileset.ListCount - 1
                If cboTileset.List(Idx) = Prj.TileSetDef(.InvIconTilesetIdx(nCurInv)).Name Then
                    cboTileset.ListIndex = Idx
                    Exit For
                End If
            Next
            txtTileIndex.Text = CStr(.InvIconTileIdx(nCurInv))
            updIconIndex.Value = .InvIconTileIdx(nCurInv)
        End If
        txtMaxQuantity.Text = CStr(.InvMaxQuantity(nCurInv))
        txtInitQuantity.Text = CStr(.InvQuantityOwned(nCurInv))
        cboQuantityDisplay.ListIndex = .InvQuantityDisplayType(nCurInv)
        txtIconCountPerRepeat.Text = CStr(.InvIconCountPerRepeat(nCurInv))
        shpBarColor.FillColor = .InvBarColor(nCurInv)
        shpBarBackground.FillColor = .InvBarBackgroundColor(nCurInv)
        If .InvBarOutlineColor(nCurInv) = -1 Then
            chkNoOutline.Value = vbChecked
        Else
            shpBarOutline.FillColor = .InvBarOutlineColor(nCurInv)
            chkNoOutline.Value = vbUnchecked
        End If
        txtBarLength.Text = .InvBarLength(nCurInv)
        txtBarThickness.Text = .InvBarThickness(nCurInv)
        txtItemX.Text = .InvDisplayX(nCurInv)
        txtItemY.Text = .InvDisplayY(nCurInv)
        If nCurInv <= 0 Then cmdPrevious.Enabled = False Else cmdPrevious.Enabled = True
        If nCurInv >= .InventoryCount - 1 Then cmdNext.Enabled = False Else cmdNext.Enabled = True
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
    End With
    Exit Sub
    
LoadInvErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub StoreInventory()
    On Error GoTo StoreErr
    
    Prj.IsDirty = True
    With Prj.GamePlayer
        If .InventoryCount <= 0 Then Exit Sub
        .InventoryItemName(nCurInv) = txtItemName.Text
        Select Case cboQuantityDisplay.ListIndex
            Case QD_HORZBARABOVE, QD_HORZBARBELOW, QD_HORZBARRIGHT, QD_VERTBARABOVE, QD_VERTBARRIGHT, QD_VERTBARLEFT
                .InvSetBarInfo nCurInv, shpBarColor.FillColor, CInt(txtBarThickness.Text), CInt(txtBarLength.Text), shpBarBackground.FillColor, IIf(chkNoOutline.Value = vbChecked, -1, shpBarOutline.FillColor)
        End Select
        If cboTileset.ListIndex < 0 Then cboTileset.ListIndex = 0
        .SetInventoryTile nCurInv, cboTileset.List(cboTileset.ListIndex), Val(txtTileIndex.Text)
        If IsNumeric(txtIconCountPerRepeat.Text) Then
            .InvIconCountPerRepeat(nCurInv) = CInt(txtIconCountPerRepeat.Text)
        End If
        If IsNumeric(txtMaxQuantity.Text) Then
            .InvMaxQuantity(nCurInv) = CInt(txtMaxQuantity.Text)
        End If
        If IsNumeric(txtInitQuantity.Text) Then
            .InvQuantityOwned(nCurInv) = CInt(txtInitQuantity.Text)
        End If
        .InvQuantityDisplayType(nCurInv) = cboQuantityDisplay.ListIndex
        If IsNumeric(txtItemX.Text) And IsNumeric(txtItemY.Text) Then
            .InvMove nCurInv, CInt(txtItemX.Text), CInt(txtItemY.Text)
        End If
    End With
    Exit Sub

StoreErr:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "PlayerSettings", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub txtBarLength_Change()
    On Error Resume Next
    If updBarLength.Value <> CInt(txtBarLength.Text) Then
        updBarLength.Value = CInt(txtBarLength.Text)
    End If
End Sub

Private Sub txtBarThickness_Change()
    On Error Resume Next
    If updBarThickness.Value <> CInt(txtBarThickness.Text) Then
        updBarThickness.Value = CInt(txtBarThickness.Text)
    End If
End Sub

Private Sub txtItemX_Change()
    On Error Resume Next
    If updItemX.Value <> CInt(txtItemX.Text) Then
        updItemX.Value = CInt(txtItemX.Text)
    End If
End Sub

Private Sub txtItemY_Change()
    On Error Resume Next
    If updItemY.Value <> CInt(txtItemY.Text) Then
        updItemY.Value = CInt(txtItemY.Text)
    End If
End Sub

Private Sub updIconIndex_Change()
    Set imgTilePreview.Picture = ExtractLocalTile(updIconIndex.Value)
End Sub

Private Function ExtractLocalTile(ByVal Index As Integer) As StdPicture
    Dim TSCols As Integer
    Dim TSRows As Integer
    
    On Error GoTo ExtractErr
    
    If cboTileset.ListIndex <= 0 Then Exit Function
    
    With Prj.TileSetDef(cboTileset.List(cboTileset.ListIndex))
        If .Image Is Nothing Then
            .Load
        End If
        If .Image Is Nothing Then Exit Function
        TSCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight
        If Index < TSRows * TSCols Then
            Set ExtractLocalTile = ExtractTile(.Image, .TileWidth * (Index Mod TSCols), .TileHeight * (Index \ TSCols), .TileWidth, .TileHeight)
        Else
            MsgBox "Tile index out of bounds", vbExclamation, "ExtractLocalTile"
        End If
    End With
    Exit Function

ExtractErr:
    MsgBox Err.Description, vbExclamation
End Function
