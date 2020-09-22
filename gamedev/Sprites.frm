VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSprites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprites and Paths"
   ClientHeight    =   6495
   ClientLeft      =   450
   ClientTop       =   330
   ClientWidth     =   7860
   HelpContextID   =   113
   Icon            =   "Sprites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   524
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeleteTemplate 
      Caption         =   "Delete Te&mplate"
      Height          =   495
      Left            =   6600
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cboTemplate 
      Height          =   315
      Left            =   4920
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadSprite 
      Caption         =   "&Load Sprite"
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtOffsetY 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtOffsetX 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdOffsetPath 
      Caption         =   "Offset Path"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkInstance 
      Caption         =   "Initial instance"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      ToolTipText     =   "Create one active instance using this definition at startup"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Timer tmrPreview 
      Interval        =   20
      Left            =   2640
      Top             =   2640
   End
   Begin VB.TextBox txtSprName 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdSprite 
      Caption         =   "U&pdate Sprite"
      Height          =   375
      Left            =   6600
      TabIndex        =   23
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelSprite 
      Caption         =   "Delete &Sprite"
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreateSprite 
      Caption         =   "&Create Sprite"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstSprites 
      Height          =   1035
      ItemData        =   "Sprites.frx":0442
      Left            =   4080
      List            =   "Sprites.frx":0444
      TabIndex        =   14
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelPoint 
      Caption         =   "Dele&te Point"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdRenPath 
      Caption         =   "&Rename Path"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstPoints 
      Height          =   645
      ItemData        =   "Sprites.frx":0446
      Left            =   1200
      List            =   "Sprites.frx":0448
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelPath 
      Caption         =   "&Delete Path"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "Sprites.frx":044A
      Left            =   120
      List            =   "Sprites.frx":044C
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame fraFrames 
      BorderStyle     =   0  'None
      Height          =   2655
      HelpContextID   =   114
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CheckBox chkAccelStates 
         Caption         =   "Separate states for accelerating and drifting"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   120
         Width           =   3615
      End
      Begin VB.ComboBox cboStates 
         Height          =   315
         ItemData        =   "Sprites.frx":044E
         Left            =   1560
         List            =   "Sprites.frx":045E
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   120
         Width           =   2175
      End
      Begin VB.VScrollBar vscrollTileset 
         Height          =   1335
         LargeChange     =   40
         Left            =   7320
         SmallChange     =   5
         TabIndex        =   44
         Top             =   1320
         Width           =   255
      End
      Begin VB.ComboBox cboSprCurState 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   480
         Width           =   2175
      End
      Begin VB.PictureBox picTileset 
         Height          =   1335
         Left            =   3840
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   229
         TabIndex        =   43
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox cboTileset 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   2640
         TabIndex        =   40
         Top             =   1320
         Width           =   1095
      End
      Begin VB.PictureBox picPreview 
         Height          =   960
         Left            =   1560
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   38
         Top             =   1560
         Width           =   960
      End
      Begin VB.CommandButton cmd36StateCopy 
         Caption         =   "<36-State<"
         Height          =   375
         Left            =   2640
         TabIndex        =   41
         ToolTipText     =   "Appends 36 tiles to the end of 36 states, beginning with selected tile"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear &All"
         Height          =   375
         Left            =   2640
         TabIndex        =   42
         Top             =   2280
         Width           =   1095
      End
      Begin MSComCtlLib.Slider sliderSpeed 
         Height          =   375
         Left            =   1560
         TabIndex        =   36
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   327682
         Max             =   25
      End
      Begin VB.Label lblStates 
         BackStyle       =   0  'Transparent
         Caption         =   "States:"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblSpriteState 
         BackStyle       =   0  'Transparent
         Caption         =   "Editing sprite state:"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAnimSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Animation Speed:"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblTileset 
         BackStyle       =   0  'Transparent
         Caption         =   "Tileset:"
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label lblPreview 
         BackStyle       =   0  'Transparent
         Caption         =   "State preview:"
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblAnimHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "(Drag tiles into preview)"
         Height          =   615
         Left            =   0
         TabIndex        =   39
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Frame fraMotion 
      BorderStyle     =   0  'None
      Height          =   2655
      HelpContextID   =   115
      Left            =   120
      TabIndex        =   45
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin MSComCtlLib.Slider sliderInertia 
         Height          =   375
         Left            =   1440
         TabIndex        =   59
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   327682
         Max             =   100
         TickFrequency   =   2
      End
      Begin MSComCtlLib.Slider sliderGravity 
         Height          =   375
         Left            =   1440
         TabIndex        =   54
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   327682
         Min             =   -10
         TickFrequency   =   10
      End
      Begin VB.CheckBox chkUpReqSolid 
         Caption         =   "Up requires solid"
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox cboSolid 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   120
         Width           =   2655
      End
      Begin VB.ComboBox cboControl 
         Height          =   315
         ItemData        =   "Sprites.frx":048E
         Left            =   1080
         List            =   "Sprites.frx":04BD
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtlLib.Slider sliderMoveSpeed 
         Height          =   375
         Left            =   1440
         TabIndex        =   52
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   327682
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblGravNone 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblDown 
         BackStyle       =   0  'Transparent
         Caption         =   "Down"
         Height          =   255
         Left            =   3600
         TabIndex        =   57
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblUp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Up"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblInertia 
         BackStyle       =   0  'Transparent
         Caption         =   "Inertia:"
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblGravPow 
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity power:"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSolid 
         BackStyle       =   0  'Transparent
         Caption         =   "Solid tiles:"
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblControl 
         BackStyle       =   0  'Transparent
         Caption         =   "Controlled by:"
         Height          =   255
         Left            =   0
         TabIndex        =   49
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblMoveSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Speed:"
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame fraCollisions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      HelpContextID   =   116
      Left            =   120
      TabIndex        =   60
      Top             =   3720
      Width           =   7575
      Begin VB.CommandButton cmdSelColl 
         Caption         =   "Select All"
         Height          =   375
         Left            =   6600
         TabIndex        =   80
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdClearColl 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   6600
         TabIndex        =   79
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdDefineCollisions 
         Caption         =   "Define..."
         Height          =   375
         Left            =   6600
         TabIndex        =   78
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame fraCollMember 
         Caption         =   "Collision Class Membership"
         Height          =   2415
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   6375
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 16"
            Height          =   255
            Index           =   15
            Left            =   3240
            TabIndex        =   77
            Top             =   2040
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 15"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   76
            Top             =   2040
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 14"
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   75
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 13"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   74
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 12"
            Height          =   255
            Index           =   11
            Left            =   3240
            TabIndex        =   73
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 11"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   72
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 10"
            Height          =   255
            Index           =   9
            Left            =   3240
            TabIndex        =   71
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 9"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 8"
            Height          =   255
            Index           =   7
            Left            =   3240
            TabIndex        =   69
            Top             =   1080
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 7"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   68
            Top             =   1080
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 6"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   67
            Top             =   840
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 5"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 4"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   65
            Top             =   600
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 2"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   63
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox chkCollClass 
            Caption         =   "Class 1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   3015
         End
      End
   End
   Begin MSComCtlLib.TabStrip tabSprite 
      Height          =   3135
      Left            =   0
      TabIndex        =   26
      Top             =   3360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Frames"
            Key             =   "Frames"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Motion"
            Key             =   "Motion"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Collisions"
            Key             =   "Collisions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTemplate 
      BackStyle       =   0  'Transparent
      Caption         =   "Template:"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblComma 
      Alignment       =   2  'Center
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblOffBy 
      Caption         =   "by:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblSprName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblSprites 
      BackStyle       =   0  'Transparent
      Caption         =   "Sprite definitions:"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblPathInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Path applies to map ... layer ..."
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblPointCount 
      BackStyle       =   0  'Transparent
      Caption         =   "(Point total)"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblPathPts 
      BackStyle       =   0  'Transparent
      Caption         =   "Path points:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblPaths 
      BackStyle       =   0  'Transparent
      Caption         =   "Paths:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSprites"
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
' File: Sprites.frm - Sprites and Paths Dialog
'
'======================================================================

Option Explicit

Dim SelTSDef As TileSetDef
Dim NewSpr As SpriteDef
Dim NewTpl As SpriteTemplate
Dim Highlighted As New TileGroup
Dim DragPic As StdPicture
Dim DragState As Integer
Dim StartDragPt As POINTAPI
Dim CFTiles As Long
Dim nPreviewRun As Integer ' -1 = Restart
Dim bLoading As Boolean

Private Sub cboSprCurState_Change()
   Dim I As Integer
   
   If (cboSprCurState.ListIndex >= 0) Then
      If Not (NewTpl.StateTilesetDef(cboSprCurState.ListIndex) Is Nothing) Then
         For I = 0 To Prj.TileSetDefCount - 1
            If cboTileset.List(I) = NewTpl.StateTilesetDef(cboSprCurState.ListIndex).Name Then
               cboTileset.ListIndex = I
               Exit For
            End If
         Next
      End If
   End If
   
End Sub

Private Sub cboSprCurState_Click()
   cboSprCurState_Change
End Sub

Private Sub cboStates_Change()
    LoadStateList
End Sub

Private Sub cboStates_Click()
   cboStates_Change
End Sub

Public Sub LoadTemplate(Tpl As SpriteTemplate)
    Dim I As Integer
    Dim TSD As TileSetDef
    
    On Error GoTo LoadTplErr
    
    bLoading = True
    Set NewTpl = Tpl.Clone
    Set NewSpr.Template = NewTpl
        
    cboTemplate.Text = NewTpl.Name
    
    If Not (NewTpl.SolidInfo Is Nothing) Then
        For I = 0 To cboSolid.ListCount - 1
            If cboSolid.List(I) = NewTpl.SolidInfo.Name Then cboSolid.ListIndex = I
        Next
    Else
        cboSolid.ListIndex = -1
    End If
    
    chkAccelStates.Value = IIf(NewTpl.Flags And eTemplateFlagBits.FLAG_ACCELSTATES, vbChecked, vbUnchecked)
    chkUpReqSolid.Value = IIf(NewTpl.Flags And eTemplateFlagBits.FLAG_UPNEEDSSOLID, vbChecked, vbUnchecked)
    
    For I = 0 To cboStates.ListCount - 1
        If cboStates.ItemData(I) = NewTpl.StateType Then
            cboStates.ListIndex = I
            Exit For
        End If
    Next
    
    sliderSpeed.Value = sliderSpeed.Max - NewTpl.AnimSpeed
    sliderMoveSpeed.Value = NewTpl.MoveSpeed
    sliderGravity.Value = NewTpl.GravPow - 10
    sliderInertia.Value = NewTpl.Inertia
    For I = 0 To 15
        If NewTpl.CollClass And 2 ^ I Then chkCollClass(I).Value = vbChecked Else chkCollClass(I).Value = vbUnchecked
    Next
    
    For I = 0 To cboControl.ListCount - 1
        If cboControl.ItemData(I) = NewTpl.ControlType Then
            cboControl.ListIndex = I
            Exit For
        End If
    Next
    
    If NewTpl.StateCount < 36 Then
        Me.cmd36StateCopy.Enabled = False
    Else
        Me.cmd36StateCopy.Enabled = True
    End If
    
    nPreviewRun = -1
    bLoading = False
    
    Exit Sub
    
LoadTplErr:
    bLoading = False
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cboTemplate_Click()
    If lstPaths.ItemData(lstPaths.ListIndex) < 0 Then Exit Sub
    If Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplateExists(cboTemplate.Text) Then LoadTemplate Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplates(cboTemplate.Text)
End Sub

Private Sub cboTileset_Change()
    Set SelTSDef = GetSelTS
    PaintTileset
End Sub

Function GetSelTS() As TileSetDef
    Dim I As Integer
    
    On Error Resume Next
    
    I = cboTileset.ListIndex
    If Prj.TileSetDef(I).Name = cboTileset.List(I) Then
        Set GetSelTS = Prj.TileSetDef(I)
    Else
        FillTilesets
        cboTileset.ListIndex = I
        Set GetSelTS = Prj.TileSetDef(I)
    End If
    If Err.Number Then
        MsgBox Err.Description, vbExclamation
    End If
End Function

Private Sub cboTileset_Click()
    cboTileset_Change
End Sub

Private Sub chkAccelStates_Click()
    If Not bLoading Then LoadStateList
End Sub

Sub LoadStateList()
    Dim I As Integer
    
    If cboStates.ListIndex < 0 Then Exit Sub
    cboSprCurState.Clear
    Select Case cboStates.ItemData(cboStates.ListIndex)
        Case eStateType.STATE_SINGLE
            If chkAccelStates.Value = vbChecked Then
                cboSprCurState.AddItem "Drifting"
                cboSprCurState.AddItem "Accelerating"
                NewTpl.StateCount = 2
            Else
                cboSprCurState.AddItem "All"
                NewTpl.StateCount = 1
            End If
        Case eStateType.STATE_LEFT_RIGHT
            If chkAccelStates.Value = vbChecked Then
                cboSprCurState.AddItem "Left Drifting"
                cboSprCurState.AddItem "Right Drifting"
                cboSprCurState.AddItem "Left Accelerating"
                cboSprCurState.AddItem "Right Accelerating"
                NewTpl.StateCount = 4
            Else
                cboSprCurState.AddItem "Left"
                cboSprCurState.AddItem "Right"
                NewTpl.StateCount = 2
            End If
        Case eStateType.STATE_8_DIRECTION
            If chkAccelStates.Value = vbChecked Then
                cboSprCurState.AddItem "Up Drifting"
                cboSprCurState.AddItem "Up-Right Drifting"
                cboSprCurState.AddItem "Right Drifting"
                cboSprCurState.AddItem "Down-Right Drifting"
                cboSprCurState.AddItem "Down Drifting"
                cboSprCurState.AddItem "Down-Left Drifting"
                cboSprCurState.AddItem "Left Drifting"
                cboSprCurState.AddItem "Up-Left Drifting"
                cboSprCurState.AddItem "Up Accelerating"
                cboSprCurState.AddItem "Up-Right Accelerating"
                cboSprCurState.AddItem "Right Accelerating"
                cboSprCurState.AddItem "Down-Right Accelerating"
                cboSprCurState.AddItem "Down Accelerating"
                cboSprCurState.AddItem "Down-Left Accelerating"
                cboSprCurState.AddItem "Left Accelerating"
                cboSprCurState.AddItem "Up-Left Accelerating"
                NewTpl.StateCount = 16
            Else
                cboSprCurState.AddItem "Up"
                cboSprCurState.AddItem "Up-Right"
                cboSprCurState.AddItem "Right"
                cboSprCurState.AddItem "Down-Right"
                cboSprCurState.AddItem "Down"
                cboSprCurState.AddItem "Down-Left"
                cboSprCurState.AddItem "Left"
                cboSprCurState.AddItem "Up-Left"
                NewTpl.StateCount = 8
            End If
        Case eStateType.STATE_36_DIRECTION
            If chkAccelStates.Value = vbChecked Then
                For I = 0 To 710 Step 10
                    Select Case I
                    Case 0
                       cboSprCurState.AddItem "Right Drifting"
                    Case 90
                       cboSprCurState.AddItem "Up Drifting"
                    Case 180
                       cboSprCurState.AddItem "Left Drifting"
                    Case 270
                       cboSprCurState.AddItem "Down Drifting"
                    Case 360
                        cboSprCurState.AddItem "Right Accelerating"
                    Case 450
                        cboSprCurState.AddItem "Up Accelerating"
                    Case 540
                        cboSprCurState.AddItem "Left Accelerating"
                    Case 630
                        cboSprCurState.AddItem "Down Accelerating"
                    Case Else
                       cboSprCurState.AddItem CStr(I Mod 360) & " Degrees " & IIf(I >= 360, "Accelerating", "Drifting")
                    End Select
                Next
                NewTpl.StateCount = 72
            Else
                For I = 0 To 350 Step 10
                    Select Case I
                    Case 0
                       cboSprCurState.AddItem "Right"
                    Case 90
                       cboSprCurState.AddItem "Up"
                    Case 180
                       cboSprCurState.AddItem "Left"
                    Case 270
                       cboSprCurState.AddItem "Down"
                    Case Else
                       cboSprCurState.AddItem CStr(I) & " Degrees"
                    End Select
                Next
                NewTpl.StateCount = 36
            End If
    End Select
    If cboSprCurState.ListCount > 0 Then cboSprCurState.ListIndex = 0
    
    If NewTpl.StateCount < 36 Then
        Me.cmd36StateCopy.Enabled = False
    Else
        Me.cmd36StateCopy.Enabled = True
    End If
    
    nPreviewRun = -1
End Sub

Private Sub chkCollClass_Click(Index As Integer)
    Dim MskVal As Integer

    If Index < 15 Then
        MskVal = 2 ^ Index
    Else
        MskVal = -32768
    End If

    If chkCollClass(Index).Value = vbChecked Then
        NewTpl.CollClass = NewTpl.CollClass Or MskVal
    Else
        NewTpl.CollClass = NewTpl.CollClass And Not MskVal
    End If

End Sub

Private Sub cmd36StateCopy_Click()
    Dim I As Integer
    Dim CurState As Integer
    Dim SelTile As Integer
    Dim TileCount As Integer
    Dim StartState As Integer
    
    If NewSpr.StateCount < 36 Then Exit Sub
    If SelTSDef Is Nothing Then Exit Sub
    
    If cboSprCurState.ListIndex < 0 Then CurState = 0 Else CurState = cboSprCurState.ListIndex
    SelTile = Highlighted.GetMember(0)
    TileCount = ScaleX(SelTSDef.Image.Width, vbHimetric, vbPixels) / SelTSDef.TileWidth * _
                ScaleY(SelTSDef.Image.Height, vbHimetric, vbPixels) / SelTSDef.TileHeight
    
    If CurState >= 36 Then StartState = 36 Else StartState = 0
    
    For I = 0 To 35
        Set NewTpl.StateTilesetDef(((CurState + I) Mod 36) + StartState) = SelTSDef
        NewTpl.AppendStateFrame ((CurState + I) Mod 36) + StartState, (SelTile + I) Mod TileCount
    Next I
End Sub

Private Sub cmdClear_Click()
    If cboSprCurState.ListIndex < 0 Then
        MsgBox "Please select a state before selecting this command."
        Exit Sub
    End If
    NewTpl.ClearState cboSprCurState.ListIndex
    nPreviewRun = -1
End Sub

Private Sub cmdClearAll_Click()
    Dim I As Integer
    
    For I = 0 To NewSpr.StateCount - 1
        NewTpl.ClearState I
    Next
End Sub

Private Sub cmdClearColl_Click()
    Dim I As Integer
    
    For I = chkCollClass.LBound To chkCollClass.UBound
        chkCollClass.Item(I).Value = vbUnchecked
    Next
End Sub

Private Sub cmdCreateSprite_Click()
    StoreAll
    If lstPaths.ListIndex < 0 Then Exit Sub
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).AddSpriteDef NewSpr.Clone
    Set NewSpr.Template = NewTpl
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).IsDirty = True
    FillSprites
    FillTemplates
End Sub

Public Sub StoreAll()
    On Error GoTo StoreErr
    Dim T As SpriteTemplate
    Dim I As Integer
    
    If lstPaths.ListIndex < 0 Then
        MsgBox "Please select a path before selecting this command.", vbExclamation
        Exit Sub
    End If
    If cboStates.ListIndex < 0 Then
        MsgBox "Please select states before selecting this command.", vbExclamation
        Exit Sub
    End If
    If cboControl.ListIndex < 0 Then
        MsgBox "Please specify controlled by before selecting this command.", vbExclamation
        Exit Sub
    End If
    If cboTemplate.Text = "" Then
        MsgBox "Please specify a name for the template of this sprite"
        Exit Sub
    End If
    
    Set T = NewTpl.Clone
    T.Name = cboTemplate.Text
    If Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplateExists(cboTemplate.Text) Then
        With Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex))
            For I = 0 To .SpriteDefCount - 1
                If .SpriteDefs(I).Template Is .SpriteTemplates(cboTemplate.Text) Then
                    Set .SpriteDefs(I).Template = T
                End If
            Next I
            Set .SpriteTemplates(cboTemplate.Text) = T
        End With
    Else
        Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).AddSpriteTemplate T
    End If
    
    NewSpr.Name = txtSprName.Text
    Set NewSpr.rPath = GetSelectedPath
    Set NewSpr.rLayer = Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).MapLayer(NewSpr.rPath.LayerName)
    NewSpr.Flags = IIf((chkInstance.Value = vbChecked), eDefFlagBits.FLAG_INSTANCE, 0)
    Set NewSpr.Template = T
    If cboSolid.ListIndex >= 0 Then
        Set NewSpr.Template.SolidInfo = Prj.SolidDefsByIndex(cboSolid.ItemData(cboSolid.ListIndex))
    Else
        Set NewSpr.Template.SolidInfo = Nothing
    End If
    NewSpr.Template.Flags = IIf((chkAccelStates.Value = vbChecked), eTemplateFlagBits.FLAG_ACCELSTATES, 0) Or _
                            IIf((chkUpReqSolid.Value = vbChecked), eTemplateFlagBits.FLAG_UPNEEDSSOLID, 0)
    NewSpr.Template.StateType = cboStates.ItemData(cboStates.ListIndex)
    NewSpr.Template.ControlType = cboControl.ItemData(cboControl.ListIndex)
    
    Exit Sub
    
StoreErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdDefineCollisions_Click()
    On Error Resume Next
    frmCollisions.Show
    If lstPaths.ListIndex >= 0 Then
        frmCollisions.cboMaps.ListIndex = lstPaths.ItemData(lstPaths.ListIndex)
    End If
End Sub

Private Sub cmdDeleteTemplate_Click()
    Dim I As Integer, J As Integer
    Dim T As SpriteTemplate
    
    If lstPaths.ItemData(lstPaths.ListIndex) < 0 Then
        MsgBox "A path must be selected (to indicate a map) before deleting a sprite template"
        Exit Sub
    End If
    
    If Not Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplateExists(cboTemplate.Text) Then
        MsgBox "Sprite template """ & cboTemplate.Text & """ not found", vbExclamation
        Exit Sub
    End If
    
    Set T = Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplates(cboTemplate.Text)
    
    For I = 0 To Prj.MapCount - 1
        With Prj.Maps(I)
            For J = 0 To .SpriteDefCount - 1
                If .SpriteDefs(J).Template Is T Then
                    MsgBox "The specified sprite template is being used by sprite definition """ & .SpriteDefs(J).Name & """ in map """ & .Name & """", vbExclamation
                    Exit Sub
                End If
            Next
        End With
    Next
    
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).RemoveSpriteTemplate T.Name
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).IsDirty = True
    
    FillTemplates
End Sub

Private Sub cmdDelPath_Click()
    Dim P As Path
    
    If lstPaths.ListIndex < 0 Then
        MsgBox "Please select a path before selecting this command.", vbExclamation
        Exit Sub
    End If
    
    Set P = GetSelectedPath
    If P.GetUsedBy Is Nothing Then
        Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).RemovePath (P.Name)
        FillPaths
    Else
        MsgBox """" & P.Name & """ is being used by sprite """ & P.GetUsedBy.Name & """"
    End If
End Sub

Private Sub cmdDelPoint_Click()
    Dim P As Path
    
    If lstPaths.ListIndex < 0 Then
        MsgBox "Please select a path before selecting this command.", vbExclamation
        Exit Sub
    End If
    If lstPoints.ListIndex < 0 Then
        MsgBox "Please select a point before selecting this command.", vbExclamation
        Exit Sub
    End If
    
    Set P = GetSelectedPath
    P.RemovePoint lstPoints.ListIndex
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).IsDirty = True
    FillPoints
End Sub

Private Sub cmdDelSprite_Click()
    If lstSprites.ListIndex < 0 Then
        MsgBox "Please select a sprite before selecting this command.", vbExclamation
        Exit Sub
    End If
    
    Prj.Maps(lstSprites.ItemData(lstSprites.ListIndex)).RemoveSpriteDef lstSprites.List(lstSprites.ListIndex)
    FillSprites
End Sub

Public Sub cmdLoadSprite_Click()
    If lstSprites.ListIndex < 0 Then Exit Sub
        
    LoadSpriteDef GetSelSpriteDef
End Sub

Private Sub cmdOffsetPath_Click()
    On Error Resume Next
    GetSelectedPath.OffsetBy CLng(txtOffsetX.Text), CLng(txtOffsetY.Text)
    If Err.Number Then
        MsgBox Err.Description
    End If
    Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).IsDirty = True
    FillPoints
    
End Sub

Private Sub cmdRenPath_Click()
    Dim S As String
    Dim P As Path

    If lstPaths.ListIndex < 0 Then
        MsgBox "Please select a path before selecting this command.", vbExclamation
        Exit Sub
    End If

    Set P = GetSelectedPath
    S = InputBox$("Enter new path name:", "Rename Path", P.Name)
    If Len(S) Then
        P.Name = S
        FillPaths
    End If
    
End Sub

Private Sub cmdSelColl_Click()
    Dim I As Integer
    
    For I = chkCollClass.LBound To chkCollClass.UBound
        chkCollClass.Item(I).Value = vbChecked
    Next
End Sub

Private Sub cmdUpdSprite_Click()
    On Error Resume Next
    
    If lstSprites.ListIndex < 0 Then
        MsgBox "Please select a sprite before selecting this command."
        Exit Sub
    End If
    
    StoreAll
    Set Prj.Maps(lstSprites.ItemData(lstSprites.ListIndex)).SpriteDefs(lstSprites.List(lstSprites.ListIndex)) = NewSpr.Clone
    Set NewSpr.Template = NewTpl
    Prj.Maps(lstSprites.ItemData(lstSprites.ListIndex)).IsDirty = True
    
    If Err.Number Then
        MsgBox Err.Description, vbExclamation
    End If
    FillSprites
    FillTemplates
End Sub

Private Sub Form_Initialize()
    CFTiles = RegisterClipboardFormat("TileGroup")
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "Sprites", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    FillPaths
    FillSprites
    FillTilesets
    tabSprite_Click
    Set NewSpr = New SpriteDef
    Set NewTpl = New SpriteTemplate
    Set NewSpr.Template = NewTpl
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nPreviewRun = 0
    SaveSetting "GameDev", "Windows", "Sprites", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub lstPaths_Click()
    Dim bFindSprite As Boolean
    Dim P As Path
    Dim S As SpriteDef
    Dim I As Integer
    Dim TSName As String
    Dim J As Integer

    On Error GoTo PathErr
    
    FillTemplates
    
    If lstPaths.ListIndex < 0 Then
        lstPoints.Clear
        Exit Sub
    End If
        
    FillPoints
    UpdatePathInfo

    Set P = GetSelectedPath
    If lstSprites.ListIndex < 0 Then
        bFindSprite = True
    Else
        If Not (GetSelSpriteDef.rPath Is P) Then
            bFindSprite = True
        End If
    End If
    
    cboSolid.Clear
    
    If lstPaths.ListIndex >= 0 Then
        TSName = Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).MapLayer(P.LayerName).TSDef.Name
        For I = 0 To Prj.SolidDefByTilesetCount(TSName) - 1
            J = Prj.SolidDefIndexByTileset(TSName, I)
            cboSolid.AddItem Prj.SolidDefsByIndex(J).Name
            cboSolid.ItemData(I) = J
        Next
    End If
    
    If bFindSprite Then
        Set S = P.GetUsedBy
        If S Is Nothing Then
            lstSprites.ListIndex = -1
            Exit Sub
        End If
        For I = 0 To lstSprites.ListCount - 1
            If Prj.Maps(lstSprites.ItemData(I)) Is S.rLayer.pMap And _
               S.Name = lstSprites.List(I) Then
                lstSprites.ListIndex = I
            End If
        Next I
    End If
    Exit Sub
    
PathErr:
    MsgBox Err.Description, vbExclamation
    
End Sub

Sub LoadSpriteDef(S As SpriteDef)
    Dim I As Integer
    Dim TSD As TileSetDef
    
    On Error GoTo LoadSprErr
    
    bLoading = True
    Set NewSpr = S.Clone
    Set NewTpl = S.Template.Clone
    Set NewSpr.Template = NewTpl
        
    If S.rPath Is Nothing Then
        lstPaths.ListIndex = -1
    Else
        For I = 0 To lstPaths.ListCount - 1
            If lstPaths.List(I) = S.rPath.Name And Prj.Maps(lstPaths.ItemData(I)).Name = S.rLayer.pMap.Name Then lstPaths.ListIndex = I
        Next I
    End If
    
    Set TSD = S.rLayer.TSDef
    
    cboTemplate.Text = NewTpl.Name
    
    txtSprName.Text = S.Name
    If Not (NewTpl.SolidInfo Is Nothing) Then
        cboSolid.ListIndex = NewTpl.SolidInfo.GetIndexByTileset(TSD.Name)
    Else
        cboSolid.ListIndex = -1
    End If
    
    chkAccelStates.Value = IIf(NewTpl.Flags And eTemplateFlagBits.FLAG_ACCELSTATES, vbChecked, vbUnchecked)
    chkUpReqSolid.Value = IIf(NewTpl.Flags And eTemplateFlagBits.FLAG_UPNEEDSSOLID, vbChecked, vbUnchecked)
    chkInstance.Value = IIf(S.Flags And eDefFlagBits.FLAG_INSTANCE, vbChecked, vbUnchecked)
    
    For I = 0 To cboStates.ListCount - 1
        If cboStates.ItemData(I) = NewTpl.StateType Then
            cboStates.ListIndex = I
            Exit For
        End If
    Next
    
    sliderSpeed.Value = sliderSpeed.Max - NewTpl.AnimSpeed
    sliderMoveSpeed.Value = NewTpl.MoveSpeed
    sliderGravity.Value = NewTpl.GravPow - 10
    sliderInertia.Value = NewTpl.Inertia
    For I = 0 To 15
        If NewTpl.CollClass And 2 ^ I Then chkCollClass(I).Value = vbChecked Else chkCollClass(I).Value = vbUnchecked
    Next
    
    For I = 0 To cboControl.ListCount - 1
        If cboControl.ItemData(I) = NewTpl.ControlType Then
            cboControl.ListIndex = I
            Exit For
        End If
    Next
    
    If NewTpl.StateCount < 36 Then
        Me.cmd36StateCopy.Enabled = False
    Else
        Me.cmd36StateCopy.Enabled = True
    End If

    nPreviewRun = -1

    bLoading = False
    Exit Sub
    
LoadSprErr:
    bLoading = False
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub FillPaths()
    Dim I As Integer
    Dim J As Integer
    
    lstPaths.Clear
    For I = 0 To Prj.MapCount - 1
        For J = 0 To Prj.Maps(I).PathCount - 1
            lstPaths.AddItem Prj.Maps(I).Paths(J).Name
            lstPaths.ItemData(lstPaths.NewIndex) = I
        Next
    Next

End Sub

Sub FillPoints()
    Dim P As Path
    Dim I As Integer

    lstPoints.Clear
    Set P = GetSelectedPath
    If P Is Nothing Then Exit Sub
    For I = 0 To P.PointCount - 1
        lstPoints.AddItem (CStr(P.PointX(I)) & ", " & CStr(P.PointY(I)))
    Next I
    lblPointCount.Caption = "(" & P.PointCount & " total)"
End Sub

Function GetSelectedPath() As Path
    If lstPaths.ListIndex < 0 Then Exit Function
    On Error Resume Next
    Set GetSelectedPath = Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).Paths(lstPaths.List(lstPaths.ListIndex))
    If Err.Number Then
        MsgBox Err.Description
    End If
End Function

Sub UpdatePathInfo()
    Dim P As Path
    
    Set P = GetSelectedPath
    If P Is Nothing Then Exit Sub
    lblPathInfo.Caption = "Applies to Map: " & Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).Name & "; Layer: " & P.LayerName
End Sub

Sub FillSprites()
    Dim I As Integer
    Dim J As Integer
    
    lstSprites.Clear
    
    For I = 0 To Prj.MapCount - 1
        For J = 0 To Prj.Maps(I).SpriteDefCount() - 1
            lstSprites.AddItem Prj.Maps(I).SpriteDefs(J).Name
            lstSprites.ItemData(lstSprites.NewIndex) = I
        Next J
    Next I
    
End Sub

Sub FillTemplates()
    Dim I As Integer
    
    cboTemplate.Clear
    If lstPaths.ItemData(lstPaths.ListIndex) < 0 Then Exit Sub
    For I = 0 To Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplateCount - 1
        cboTemplate.AddItem Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).SpriteTemplates(I).Name
    Next
    For I = 0 To 15
        chkCollClass(I).Caption = Prj.Maps(lstPaths.ItemData(lstPaths.ListIndex)).CollClassName(I)
    Next
End Sub

Function GetSelSpriteDef() As SpriteDef
    
    On Error Resume Next
    Set GetSelSpriteDef = Prj.Maps(lstSprites.ItemData(lstSprites.ListIndex)).SpriteDefs(lstSprites.List(lstSprites.ListIndex))
    If Err.Number Then
        MsgBox Err.Description, vbExclamation
    End If
    
End Function

Sub FillTilesets()
    Dim I As Integer
    
    cboTileset.Clear
    For I = 0 To Prj.TileSetDefCount - 1
        cboTileset.AddItem Prj.TileSetDef(I).Name
    Next
End Sub

Private Sub picPreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo DropErr
    If cboSprCurState.ListIndex < 0 Then
        MsgBox "Please select a state before adding to the animation.", vbExclamation
        Exit Sub
    End If
    Set NewTpl.StateTilesetDef(cboSprCurState.ListIndex) = Prj.TileSetDef(cboTileset.List(cboTileset.ListIndex))
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectCopy
        NewTpl.AppendStateFrame cboSprCurState.ListIndex, Data.GetData(CInt("&H" & Hex$(CFTiles)))(2)
    Else
        Effect = vbDropEffectNone
    End If
    DragState = 0
    
    If nPreviewRun > 0 Then
        nPreviewRun = -1 ' Restart anim
    Else
        tmrPreview.Enabled = True
    End If
    
    Exit Sub
    
DropErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub picPreview_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(CInt("&H" & Hex$(CFTiles))) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
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

Sub PaintTileset()
    Dim TSCols As Integer
    Dim TSRows As Integer
    Dim I As Integer
    Dim rcTile As RECT
    Dim YMax As Integer

    If SelTSDef Is Nothing Then Exit Sub
    
    On Error GoTo PaintErr

    With SelTSDef
        If Not .IsLoaded Then .Load
        If Not .IsLoaded Then Exit Sub
        
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
            picTileset.PaintPicture ExtractLocalTile(I, SelTSDef, Highlighted.IsMember(I)), rcTile.Left, rcTile.Top
        End If
    Next
    
    vscrollTileset.Max = YMax
    
    Exit Sub
    
PaintErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function ExtractLocalTile(ByVal Index As Integer, TSD As TileSetDef, Optional bHighlight As Boolean = False) As StdPicture
    Dim TSCols As Integer
    Dim TSRows As Integer
    
    With TSD
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
        
    With SelTSDef
        FitCols = (picTileset.ScaleWidth) \ (.TileWidth + 6)
        GetTilesetTileRect.Left = (Index Mod FitCols) * (.TileWidth + 6) + 3
        GetTilesetTileRect.Top = (Index \ FitCols) * (.TileHeight + 6) + 3
        GetTilesetTileRect.Right = GetTilesetTileRect.Left + .TileWidth - 1
        GetTilesetTileRect.Bottom = GetTilesetTileRect.Top + .TileHeight - 1
    End With
    
End Function

Private Sub sliderGravity_Change()
    NewTpl.GravPow = sliderGravity.Value + 10
End Sub

Private Sub sliderGravity_Click()
    NewTpl.GravPow = sliderGravity.Value + 10
End Sub

Public Sub sliderInertia_Change()
    NewTpl.Inertia = sliderInertia.Value
End Sub

Private Sub sliderInertia_Click()
    NewTpl.Inertia = sliderInertia.Value
End Sub

Private Sub sliderMoveSpeed_Change()
   NewTpl.MoveSpeed = sliderMoveSpeed.Value
End Sub

Private Sub sliderMoveSpeed_Click()
   NewTpl.MoveSpeed = sliderMoveSpeed.Value
End Sub

Private Sub sliderSpeed_Change()
   NewTpl.AnimSpeed = sliderSpeed.Max - sliderSpeed.Value
End Sub

Private Sub sliderSpeed_Click()
   NewTpl.AnimSpeed = sliderSpeed.Max - sliderSpeed.Value
End Sub

Private Sub tabSprite_Click()
    Select Case tabSprite.SelectedItem.Key
    Case "Frames"
        fraFrames.Visible = True
        fraMotion.Visible = False
        fraCollisions.Visible = False
    Case "Motion"
        fraFrames.Visible = False
        fraMotion.Visible = True
        fraCollisions.Visible = False
    Case "Collisions"
        fraFrames.Visible = False
        fraMotion.Visible = False
        fraCollisions.Visible = True
    End Select
End Sub

Private Sub tmrPreview_Timer()
    Dim nTimerSpeed As Long
    Dim T As Single
    Dim I As Integer
    Dim SprInst As Sprite
    Dim bSkipAnim As Boolean
    
    tmrPreview.Enabled = False
    picPreview.Cls
    If nPreviewRun > 0 Then Exit Sub
    nPreviewRun = 1
    
    T = Timer
    Do
        nTimerSpeed = nTimerSpeed + 1
        DoEvents
    Loop Until Timer - T >= 0.5

    If nPreviewRun = 0 Then Exit Sub

    nTimerSpeed = nTimerSpeed / 10

    NewTpl.AnimSpeed = sliderSpeed.Max - sliderSpeed.Value
    Set SprInst = NewSpr.MakeInstance

    Do While nPreviewRun > 0
    
        bSkipAnim = False
        'If NewSpr Is Nothing Then Exit Sub
        If cboSprCurState.ListIndex >= NewTpl.StateCount Or cboSprCurState.ListIndex < 0 Then
            bSkipAnim = True
        Else
            If SprInst.CurState <> cboSprCurState.ListIndex Then
                picPreview.Cls
                SprInst.CurState = cboSprCurState.ListIndex
                SprInst.ResetFrame
            End If
            If NewTpl.StateFrameCount(SprInst.CurState) = 0 Then
                bSkipAnim = True
            End If
        End If
         
        If bSkipAnim = True Then
            picPreview.Cls
        Else
            With SprInst.CurTSDef
                picPreview.PaintPicture ExtractLocalTile(SprInst.CurTile, SprInst.CurTSDef), (picPreview.ScaleWidth - .TileWidth) / 2, (picPreview.ScaleHeight - .TileHeight) / 2
            End With
            SprInst.AdvanceFrame 1
        End If
        
        T = 0
        Do
            T = T + 1
            DoEvents
        Loop Until T >= nTimerSpeed Or nPreviewRun <= 0
    Loop

    If nPreviewRun < 0 Then tmrPreview.Enabled = True

    Set SprInst.rDef = Nothing
    
End Sub

Private Sub vscrollTileset_Change()
   PaintTileset
End Sub

Private Sub vscrollTileset_Scroll()
   PaintTileset
End Sub

Private Function GetXYTileSetTile(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim TSCols As Integer
    Dim TSRows As Integer
    Dim I As Integer
    Dim rcTile As RECT
    
    GetXYTileSetTile = -1
    
    'If NewSpr Is Nothing Then Exit Function

    With SelTSDef
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

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button, Shift, X, Y
End Sub

Private Sub HandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Idx As Integer
    Dim H As TileGroup
    
    If Button = 1 Then
        Idx = GetXYTileSetTile(X, Y)
        Set H = Highlighted
        
        If Idx >= 0 Then
            If Shift And vbCtrlMask Then
                If H.IsMember(Idx) Then
                    H.ClearMember Idx
                Else
                    H.SetMember Idx
                    Set DragPic = ExtractLocalTile(Idx, SelTSDef)
                    DragState = 1
                    StartDragPt.X = X
                    StartDragPt.Y = Y
                End If
                PaintTileset
            Else
                If Not H.IsMember(Idx) Then
                     SelectSingle Idx
                End If
                Set DragPic = ExtractLocalTile(Idx, SelTSDef)
                DragState = 1
                StartDragPt.X = X
                StartDragPt.Y = Y
            End If
        Else
            If (Shift And vbCtrlMask) = 0 Then
                If Not H.IsEmpty Then
                    H.ClearAll
                    PaintTileset
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
