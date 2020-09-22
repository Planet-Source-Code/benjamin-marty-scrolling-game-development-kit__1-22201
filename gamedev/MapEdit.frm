VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMapEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Maps"
   ClientHeight    =   6780
   ClientLeft      =   2340
   ClientTop       =   165
   ClientWidth     =   6420
   HelpContextID   =   105
   Icon            =   "MapEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgMapFile 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "map"
      DialogTitle     =   "Specify map filename"
      Filter          =   "All Files (*.*)|*.*|Map files (*.map)|*.map"
      FilterIndex     =   2
      Flags           =   34822
   End
   Begin VB.CommandButton cmdSaveMap 
      Caption         =   "Save Map"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraMapEdit 
      Caption         =   "Map Display"
      Height          =   1575
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   270
      Width           =   1455
      Begin VB.TextBox txtDispLeft 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtDispHeight 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   1185
         Width           =   615
      End
      Begin VB.TextBox txtDispWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   870
         Width           =   615
      End
      Begin VB.TextBox txtDispTop 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   560
         Width           =   615
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Height::"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   9
         Top             =   885
         Width           =   495
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   7
         Top             =   575
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdUpdateMap 
      Caption         =   "Update Map"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      ToolTipText     =   "Updates parameters of the selected map, redefining based on displayed values"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtMapHeight 
      Height          =   285
      Left            =   3960
      TabIndex        =   17
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtMapWidth 
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdDeleteMap 
      Caption         =   "Delete Map"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtMapName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewMap 
      Caption         =   "New Map"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox lstMaps 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame fraLayers 
      BorderStyle     =   0  'None
      Height          =   3735
      HelpContextID   =   106
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   6135
      Begin VB.CommandButton cmdEditMap 
         Caption         =   "Edit"
         Height          =   375
         HelpContextID   =   107
         Left            =   4800
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox lstLayers 
         Height          =   840
         ItemData        =   "MapEdit.frx":0442
         Left            =   120
         List            =   "MapEdit.frx":0444
         TabIndex        =   27
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton cmdNewLayer 
         Caption         =   "New Layer"
         Height          =   375
         Left            =   4800
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLayerName 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cboLayerTileset 
         Height          =   315
         ItemData        =   "MapEdit.frx":0446
         Left            =   1200
         List            =   "MapEdit.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtXRate 
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtYRate 
         Height          =   285
         Left            =   2760
         TabIndex        =   39
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdMoveUp 
         Height          =   245
         Left            =   3120
         Picture         =   "MapEdit.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdMoveDown 
         Height          =   245
         Left            =   3120
         Picture         =   "MapEdit.frx":054C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   255
      End
      Begin VB.Frame fraMapEdit 
         Caption         =   "Screen Depth"
         Height          =   1095
         Index           =   0
         Left            =   3360
         TabIndex        =   40
         Top             =   1320
         Width           =   1335
         Begin VB.OptionButton opt16Bit 
            Caption         =   "16-Bit color"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt24Bit 
            Caption         =   "24-bit color"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton opt32Bit 
            Caption         =   "32-bit color"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdDeleteLayer 
         Caption         =   "Delete Layer"
         Height          =   375
         Left            =   4800
         TabIndex        =   46
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdateLayer 
         Caption         =   "Update Layer"
         Height          =   375
         Left            =   4800
         TabIndex        =   48
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "Transparent layer"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Map Layers:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Layer Name:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Tileset:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "X Scroll Rate:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Scroll Rate:"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   38
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Background"
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblMapEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreground"
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame fraPlayerInteraction 
      Caption         =   "Player Interaction"
      Height          =   3375
      HelpContextID   =   108
      Left            =   120
      TabIndex        =   52
      Top             =   3240
      Width           =   6135
      Begin VB.ComboBox cboTouchMedia 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox cboTouchCategory 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cboReaction 
         Height          =   315
         ItemData        =   "MapEdit.frx":064E
         Left            =   3240
         List            =   "MapEdit.frx":065E
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cboRelaventInventory 
         Height          =   315
         ItemData        =   "MapEdit.frx":06CD
         Left            =   3240
         List            =   "MapEdit.frx":06CF
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   960
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton optFirstTouch 
            Caption         =   "Initially"
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optContinuousTouch 
            Caption         =   "Continuously"
            Height          =   255
            Left            =   1200
            TabIndex        =   57
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   5775
         Begin VB.OptionButton optRemoveIfUsed 
            Caption         =   "Remove tile if inventory OK"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   2655
         End
         Begin VB.OptionButton optDontRemove 
            Caption         =   "Don't remove"
            Height          =   255
            Left            =   2640
            TabIndex        =   63
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optAlwaysRemove 
            Caption         =   "Always Remove"
            Height          =   255
            Left            =   4200
            TabIndex        =   64
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.TextBox txtReplaceTile 
         Height          =   285
         Left            =   3240
         TabIndex        =   66
         Top             =   1680
         Width           =   1020
      End
      Begin VB.CheckBox chkRaiseEvent 
         Caption         =   "Raise an event for this interaction"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton cmdFirstMI 
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
         TabIndex        =   71
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevMI 
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
         TabIndex        =   72
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdNextMI 
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
         TabIndex        =   73
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdLastMI 
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
         TabIndex        =   74
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdNewMI 
         Caption         =   "&New"
         Height          =   375
         Left            =   3240
         TabIndex        =   75
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteMI 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4200
         TabIndex        =   76
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdateMI 
         Caption         =   "&Update"
         Height          =   375
         Left            =   5160
         TabIndex        =   77
         Top             =   2880
         Width           =   855
      End
      Begin MSComCtl2.UpDown updReplaceTile 
         Height          =   285
         Left            =   4261
         TabIndex        =   67
         Top             =   1680
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtReplaceTile"
         BuddyDispid     =   196651
         OrigLeft        =   4440
         OrigTop         =   1920
         OrigRight       =   4635
         OrigBottom      =   2295
         Max             =   255
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTouchMedia 
         Caption         =   "Play media clip:"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblTouchCategory 
         Caption         =   "When player touches a tile in category:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblInventoryItem 
         Caption         =   "Relevant inventory item:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblReplaceTile 
         Caption         =   "Replace removed tile with:"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Image imgReplaceTile 
         Height          =   975
         Left            =   5040
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame fraPlayerSprite 
      BorderStyle     =   0  'None
      Height          =   375
      HelpContextID   =   108
      Left            =   120
      TabIndex        =   49
      Top             =   2880
      Width           =   6135
      Begin VB.ComboBox cboSprites 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   60
         Width           =   2775
      End
      Begin VB.Label lblPlayerSprite 
         Caption         =   "Player sprite (""Initial Instance"" required):"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   60
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpecialFunctions 
      Caption         =   "Special Functions"
      Height          =   3735
      HelpContextID   =   109
      Left            =   120
      TabIndex        =   78
      Top             =   2880
      Width           =   6135
      Begin VB.CommandButton cmdCalc 
         Height          =   360
         Left            =   4560
         Picture         =   "MapEdit.frx":06D1
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Run Windows Calculator"
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpdateFunc 
         Caption         =   "&Update"
         Height          =   375
         Left            =   5040
         TabIndex        =   145
         Top             =   3240
         Width           =   975
      End
      Begin VB.ListBox lstSpecialFunctions 
         Height          =   2790
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdDeleteFunc 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   143
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame fraFuncAct 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   1920
         TabIndex        =   81
         Top             =   600
         Width           =   3975
         Begin VB.Frame fraNotOnLayer 
            BorderStyle     =   0  'None
            Caption         =   "fraNotOnLayer"
            Height          =   1935
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Visible         =   0   'False
            Width           =   3975
            Begin VB.Label lblNotOnLayer 
               Alignment       =   2  'Center
               Caption         =   "The selected function is not on the same layer as the player sprite and cannot be activated by the player."
               Height          =   855
               Left            =   240
               TabIndex        =   155
               Top             =   600
               Width           =   3615
            End
         End
         Begin VB.CheckBox chkActOnStartup 
            Caption         =   "Activate once before gameplay begins"
            Height          =   255
            Left            =   0
            TabIndex        =   95
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CheckBox chkFuncUseInvRemove 
            Caption         =   "Remove item(s) from inventory on use"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1680
            Width           =   3735
         End
         Begin MSComCtl2.UpDown updFuncUseInv 
            Height          =   285
            Left            =   1740
            TabIndex        =   93
            Top             =   1320
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtFuncUseInvCount"
            BuddyDispid     =   196676
            OrigLeft        =   1800
            OrigTop         =   1200
            OrigRight       =   1995
            OrigBottom      =   1500
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtFuncUseInvCount 
            Height          =   285
            Left            =   1200
            TabIndex        =   92
            Top             =   1320
            Width           =   540
         End
         Begin VB.ComboBox cboFuncUseInv 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkFuncUseInventory 
            Caption         =   "Uses inventory item"
            Height          =   255
            Left            =   0
            TabIndex        =   88
            Top             =   690
            Width           =   1935
         End
         Begin VB.CheckBox chkFuncUp 
            Caption         =   "Going Up"
            Height          =   255
            Left            =   840
            TabIndex        =   83
            Top             =   120
            Width           =   1095
         End
         Begin VB.CheckBox chkFuncButton 
            Caption         =   "Button"
            Height          =   255
            Left            =   2040
            TabIndex        =   84
            Top             =   120
            Width           =   855
         End
         Begin VB.CheckBox chkFuncDown 
            Caption         =   "Down"
            Height          =   255
            Left            =   3000
            TabIndex        =   85
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox chkRemoveFunc 
            Caption         =   "Remove after use"
            Height          =   255
            Left            =   840
            TabIndex        =   86
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkFuncInitial 
            Caption         =   "Initial touch only"
            Height          =   255
            Left            =   2520
            TabIndex        =   87
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblFuncUseInvItm 
            Caption         =   "Which item:"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblFuncUseInvCount 
            Caption         =   "How many:"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblOnlyWhen 
            Caption         =   "Only when:"
            Height          =   255
            Left            =   0
            TabIndex        =   82
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraFuncDef 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   1920
         TabIndex        =   96
         Top             =   600
         Width           =   3975
         Begin VB.ComboBox cboFunction 
            Height          =   315
            ItemData        =   "MapEdit.frx":0A13
            Left            =   840
            List            =   "MapEdit.frx":0A2F
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   120
            Width           =   3135
         End
         Begin VB.Frame fraFuncMessage 
            Caption         =   "Message"
            Height          =   1935
            Left            =   0
            TabIndex        =   99
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            Begin VB.TextBox txtFuncMessage 
               Height          =   1095
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   100
               Top             =   240
               Width           =   3735
            End
            Begin VB.Label lblDisplayPicture 
               BackStyle       =   0  'Transparent
               Caption         =   "To display a picture instead of text, begin the message with #PIC followed by the file path."
               Height          =   495
               Left            =   120
               TabIndex        =   146
               Top             =   1440
               Width           =   3735
            End
         End
         Begin VB.Frame fraSwitchSprite 
            Caption         =   "Switch to Sprite"
            Height          =   1935
            Left            =   0
            TabIndex        =   101
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            Begin VB.ListBox lstSwitchSprite 
               Height          =   1425
               Left            =   120
               TabIndex        =   102
               Top             =   360
               Width           =   1695
            End
            Begin VB.CheckBox chkSSSwapControl 
               Caption         =   "Swap ""Controlled by"""
               Height          =   255
               Left            =   1920
               TabIndex        =   103
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox chkSSNewInstance 
               Caption         =   "New instance"
               Height          =   255
               Left            =   1920
               TabIndex        =   104
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox chkSSDeleteOld 
               Caption         =   "Delete old sprite"
               Height          =   255
               Left            =   1920
               TabIndex        =   105
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CheckBox chkSSSameLocation 
               Caption         =   "Same location as old"
               Height          =   255
               Left            =   1920
               TabIndex        =   106
               Top             =   1440
               Width           =   1935
            End
         End
         Begin VB.Frame fraSwitchMap 
            Caption         =   "Switch to Map"
            Height          =   1935
            Left            =   0
            TabIndex        =   107
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            Begin VB.ListBox lstSwitchMap 
               Height          =   1425
               Left            =   120
               TabIndex        =   108
               Top             =   360
               Width           =   1815
            End
            Begin VB.ComboBox cboSMSprite 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   600
               Width           =   1815
            End
            Begin VB.CheckBox chkSMPosOverride 
               Caption         =   "Set start position"
               Height          =   255
               Left            =   2040
               TabIndex        =   111
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtSMX 
               Height          =   285
               Left            =   2280
               TabIndex        =   113
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox txtSMY 
               Height          =   285
               Left            =   3240
               TabIndex        =   115
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lblSMSprite 
               Caption         =   "As this sprite:"
               Height          =   255
               Left            =   2040
               TabIndex        =   109
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label lblSMX 
               Caption         =   "X:"
               Height          =   255
               Left            =   2040
               TabIndex        =   112
               Top             =   1320
               Width           =   255
            End
            Begin VB.Label lblSMY 
               Caption         =   "Y:"
               Height          =   255
               Left            =   3000
               TabIndex        =   114
               Top             =   1320
               Width           =   255
            End
            Begin VB.Label lblSMInstruct 
               Caption         =   "(Specify in pixels)"
               Height          =   255
               Left            =   2280
               TabIndex        =   116
               Top             =   1620
               Width           =   1455
            End
         End
         Begin VB.Frame fraTeleport 
            Caption         =   "Teleport Location"
            Height          =   1935
            Left            =   0
            TabIndex        =   117
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            Begin VB.PictureBox picTeleport 
               AutoRedraw      =   -1  'True
               Height          =   1575
               Left            =   1920
               ScaleHeight     =   101
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   125
               TabIndex        =   123
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtTeleportX 
               Height          =   285
               Left            =   1080
               TabIndex        =   119
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtTeleportY 
               Height          =   285
               Left            =   1080
               TabIndex        =   121
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox chkTeleportOffset 
               Caption         =   "Offset from current"
               Height          =   255
               Left            =   120
               TabIndex        =   122
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lblTeleportX 
               Caption         =   "X (in pixels):"
               Height          =   255
               Left            =   120
               TabIndex        =   118
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblTeleportY 
               Caption         =   "Y (in pixels):"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame fraAlterMap 
            Caption         =   "Alter Map"
            Height          =   1935
            Left            =   0
            TabIndex        =   124
            Top             =   480
            Width           =   3975
            Begin VB.ListBox lstAlterMapFunc 
               Height          =   1035
               Left            =   120
               TabIndex        =   126
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtCopyToX 
               Height          =   285
               Left            =   2400
               TabIndex        =   129
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtCopyToY 
               Height          =   285
               Left            =   3360
               TabIndex        =   131
               Top             =   720
               Width           =   495
            End
            Begin VB.ComboBox cboPickCoord 
               Height          =   315
               ItemData        =   "MapEdit.frx":0ACB
               Left            =   2160
               List            =   "MapEdit.frx":0ADB
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label lblCopyFrom 
               Caption         =   "Copy tiles from function:"
               Height          =   255
               Left            =   120
               TabIndex        =   125
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label lblCopyTo 
               Caption         =   "Copy to tile coords:"
               Height          =   255
               Left            =   2280
               TabIndex        =   127
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lblCopyToX 
               Caption         =   "X:"
               Height          =   255
               Left            =   2160
               TabIndex        =   128
               Top             =   720
               Width           =   255
            End
            Begin VB.Label lblCopyToY 
               Caption         =   "Y:"
               Height          =   255
               Left            =   3120
               TabIndex        =   130
               Top             =   720
               Width           =   255
            End
            Begin VB.Label lblPickCoord 
               Caption         =   "Pick coordinates:"
               Height          =   255
               Left            =   2160
               TabIndex        =   132
               Top             =   1080
               Width           =   1695
            End
         End
         Begin VB.Frame fraCreateSprite 
            Caption         =   "Create Sprite"
            Height          =   1935
            Left            =   0
            TabIndex        =   134
            Top             =   480
            Width           =   3975
            Begin VB.TextBox txtCreateMaxCount 
               Height          =   285
               Left            =   3240
               TabIndex        =   148
               Top             =   1440
               Width           =   615
            End
            Begin VB.ListBox lstCreateSprite 
               Height          =   1425
               Left            =   120
               TabIndex        =   135
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtCreateSpriteX 
               Height          =   285
               Left            =   2160
               TabIndex        =   138
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtCreateSpriteY 
               Height          =   285
               Left            =   3240
               TabIndex        =   140
               Top             =   720
               Width           =   615
            End
            Begin VB.OptionButton optCreatePosAbsolute 
               Caption         =   "Absolute"
               Height          =   255
               Left            =   1920
               TabIndex        =   141
               Top             =   1080
               Width           =   975
            End
            Begin VB.OptionButton optCreatePosPlrRel 
               Caption         =   "Relative"
               Height          =   255
               Left            =   2880
               TabIndex        =   142
               Top             =   1080
               Width           =   975
            End
            Begin VB.CheckBox chkCreatePos 
               Caption         =   "Set start position (pixel)"
               Height          =   255
               Left            =   1920
               TabIndex        =   136
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label lblCreateCount 
               Caption         =   "Maximum count:"
               Height          =   255
               Left            =   1920
               TabIndex        =   147
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lblCreateSpriteX 
               Caption         =   "X:"
               Height          =   255
               Left            =   1920
               TabIndex        =   137
               Top             =   720
               Width           =   255
            End
            Begin VB.Label lblCreateSpriteY 
               Caption         =   "Y:"
               Height          =   255
               Left            =   3000
               TabIndex        =   139
               Top             =   720
               Width           =   255
            End
         End
         Begin VB.Frame fraDeleteSprite 
            Caption         =   "Delete Sprites"
            Height          =   1935
            Left            =   0
            TabIndex        =   149
            Top             =   480
            Width           =   3975
            Begin VB.CheckBox chkDeleteMany 
               Caption         =   "Delete all above min."
               Height          =   255
               Left            =   1920
               TabIndex        =   153
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtDeleteCount 
               Height          =   285
               Left            =   3240
               TabIndex        =   152
               Top             =   360
               Width           =   615
            End
            Begin VB.ListBox lstDeleteSprite 
               Height          =   1425
               Left            =   120
               TabIndex        =   150
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lblDeleteCount 
               Caption         =   "Minimum count:"
               Height          =   255
               Left            =   1920
               TabIndex        =   151
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Label lblFunction 
            BackStyle       =   0  'Transparent
            Caption         =   "Function:"
            Height          =   255
            Left            =   0
            TabIndex        =   97
            Top             =   120
            Width           =   855
         End
      End
      Begin MSComctlLib.TabStrip tabFunctions 
         Height          =   2895
         Left            =   1800
         TabIndex        =   80
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5106
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Action Parameters"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Effect"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tabMaps 
      Height          =   4215
      Left            =   0
      TabIndex        =   24
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Layers"
            Key             =   "Layers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player Interaction"
            Key             =   "PlayerInteraction"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Special Functions"
            Key             =   "SpecialFunctions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMapPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   2235
      Width           =   3495
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   18
      Top             =   2235
      Width           =   1095
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Map size in pixels:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Map Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblMapEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Maps:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MapEdit.frm - Maps and Layers Dialog
'
'======================================================================

Option Explicit

Public nCurInt As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cboFunction_Change()
    LoadFunc True
End Sub

Private Sub cboFunction_Click()
    LoadFunc True
End Sub

Private Sub cboPickCoord_Change()
    cboPickCoord_Click
End Sub

Private Sub cboPickCoord_Click()
    Dim OtherFunc As SpecialFunction
    Dim ThisFunc As SpecialFunction

    On Error GoTo PickCoordErr
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstSpecialFunctions.ListIndex < 0 Then Exit Sub
    If lstAlterMapFunc.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        If Not .SpecialExists(lstSpecialFunctions.Text) Then Exit Sub
        If Not .SpecialExists(lstAlterMapFunc.Text) Then Exit Sub
        Set OtherFunc = .Specials(lstAlterMapFunc.Text)
        Set ThisFunc = .Specials(lstSpecialFunctions.Text)
        Select Case cboPickCoord.ListIndex
        Case 0  ' Match top left
            txtCopyToX.Text = CStr(ThisFunc.TileLeft)
            txtCopyToY.Text = CStr(ThisFunc.TileTop)
        Case 1 ' Match top right
            txtCopyToX.Text = CStr(ThisFunc.TileRight - (OtherFunc.TileRight - OtherFunc.TileLeft))
            txtCopyToY.Text = CStr(ThisFunc.TileTop)
        Case 2 ' Match bottom left
            txtCopyToX.Text = CStr(ThisFunc.TileLeft)
            txtCopyToY.Text = CStr(ThisFunc.TileBottom - (OtherFunc.TileBottom - OtherFunc.TileTop))
        Case 3 ' Match bottom right
            txtCopyToX.Text = CStr(ThisFunc.TileRight - (OtherFunc.TileRight - OtherFunc.TileLeft))
            txtCopyToY.Text = CStr(ThisFunc.TileBottom - (OtherFunc.TileBottom - OtherFunc.TileTop))
        End Select
    End With
    Exit Sub
    
PickCoordErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cboSprites_Change()
    FillCategories
    LoadReplaceTile
    If lstMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        .PlayerSpriteName = cboSprites.List(cboSprites.ListIndex)
    End With
End Sub

Private Sub cboSprites_Click()
    cboSprites_Change
End Sub

Private Sub chkCreatePos_Click()
    Dim bEnable As Boolean
    
    If chkCreatePos.Value = vbChecked Then
        bEnable = True
    Else
        bEnable = False
    End If
    
    lblCreateSpriteX.Enabled = bEnable
    txtCreateSpriteX.Enabled = bEnable
    lblCreateSpriteY.Enabled = bEnable
    txtCreateSpriteY.Enabled = bEnable
    optCreatePosAbsolute.Enabled = bEnable
    optCreatePosPlrRel.Enabled = bEnable
    
End Sub

Private Sub chkFuncUseInventory_Click()
    Dim bEnable As Boolean
    bEnable = (chkFuncUseInventory.Value = vbChecked)
    
    lblFuncUseInvItm.Enabled = bEnable
    cboFuncUseInv.Enabled = bEnable
    lblFuncUseInvCount.Enabled = bEnable
    txtFuncUseInvCount.Enabled = bEnable
    updFuncUseInv.Enabled = bEnable
    chkFuncUseInvRemove.Enabled = bEnable
    If bEnable Then
        If Val(txtFuncUseInvCount.Text) <= 0 Then txtFuncUseInvCount = 1
    Else
        txtFuncUseInvCount.Text = "0"
    End If
    
        
End Sub

Private Sub chkTeleportOffset_Click()
    DrawTeleportPreview
End Sub

Private Sub cmdCalc_Click()
    On Error Resume Next
    Shell "Calc.exe", vbNormalFocus
End Sub

Private Sub cmdDeleteFunc_Click()
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstSpecialFunctions.ListIndex < 0 Then
        MsgBox "Please select a special function to delete first."
        Exit Sub
    End If
    Prj.Maps(lstMaps.ListIndex).RemoveSpecial lstSpecialFunctions.List(lstSpecialFunctions.ListIndex)
    FillSpecialFunctions
End Sub

Private Sub cmdDeleteLayer_Click()
    If lstMaps.ListIndex < 0 Or lstLayers.ListIndex < 0 Then
        MsgBox "Please select a map and a layer before selecting this command", vbExclamation
        Exit Sub
    End If
    Prj.Maps(lstMaps.ListIndex).RemoveLayer lstLayers.List(lstLayers.ListIndex)
    UpdateLayerList
End Sub

Private Sub cmdDeleteMap_Click()
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map before selecting this command", vbExclamation
        Exit Sub
    End If
    Prj.RemoveMap lstMaps.List(lstMaps.ListIndex)
    UpdateMapList
End Sub

Private Sub cmdEditMap_Click()
    Dim M As New MapEdit
    
    On Error GoTo EditErr
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstLayers.ListIndex < 0 Then
        MsgBox "A map and a layer must be selected", vbExclamation
        Exit Sub
    End If
    With Prj.Maps(lstMaps.ListIndex)
        If .ViewLeft + .ViewWidth > 640 Or .ViewLeft < 0 Or .ViewTop + .ViewHeight > 480 Or .ViewTop < 0 Or .ViewHeight <= 0 Or .ViewWidth <= 0 Then
            MsgBox "The display parameters for the map are out of range.", vbExclamation
            Exit Sub
        End If
    End With
    Set M.Disp = New BMDXDisplay
    On Error Resume Next
    M.Disp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
    On Error GoTo EditErr
    Prj.TriggerEditMap M
    If Not GameHost Is Nothing Then
        GameHost.RunStartScript
        If GameHost.CheckForError Then Exit Sub
    End If
    M.Disp.OpenEx , , IIf(opt24Bit.Value, 24, IIf(opt32Bit.Value, 32, 16))
    Set CurDisp = M.Disp
    M.Edit Prj.Maps(lstMaps.ListIndex), lstLayers.ListIndex
    Set CurDisp = Nothing
    Set M = Nothing
    Exit Sub

EditErr:
    Set M = Nothing
    MsgBox Err.Description
End Sub

Private Sub cmdMoveDown_Click()
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map and a layer before selecting this command"
        Exit Sub
    End If
    
    If lstLayers.ListIndex >= lstLayers.ListCount - 1 Then
        MsgBox "Please select a layer (above the bottom) before selecting this command", vbExclamation
        Exit Sub
    End If
    
    With Prj.Maps(lstMaps.ListIndex)
        .ShiftLayer lstLayers.ListIndex, 1
    End With
    
    UpdateLayerList
End Sub

Private Sub cmdMoveUp_Click()
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map and a layer before selecting this command"
        Exit Sub
    End If
    
    If lstLayers.ListIndex <= 0 Then
        MsgBox "Please select a layer (below the top) before selecting this command", vbExclamation
        Exit Sub
    End If
    
    With Prj.Maps(lstMaps.ListIndex)
        .ShiftLayer lstLayers.ListIndex, -1
    End With
    
    UpdateLayerList
End Sub

Private Sub cmdNewLayer_Click()
    On Error GoTo LayerCreateErr
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map to which the layer should be added before selecting this command", vbExclamation
        Exit Sub
    End If
    If Prj.Maps(lstMaps.ListIndex).LayerExists(txtLayerName.Text) Or txtLayerName.Text = "" Then
        MsgBox "Please enter a new layer name before creating it.", vbExclamation
        Exit Sub
    End If
    
    Prj.Maps(lstMaps.ListIndex).AddLayer txtLayerName.Text, cboLayerTileset.Text, CSng(txtXRate.Text), CSng(txtYRate.Text), (chkTransparent.Value = vbChecked)
    UpdateLayerList
    Exit Sub

LayerCreateErr:
    MsgBox Err.Description
End Sub

Private Sub cmdNewMap_Click()
    Dim M As Map
    
    On Error GoTo MapCreateErr
    If Prj.MapExists(txtMapName.Text) Or txtMapName.Text = "" Then
        MsgBox "Please enter a new name for the map before adding it", vbExclamation
        Exit Sub
    End If
    Set M = New Map
    M.Name = txtMapName.Text
    M.MapWidth = CLng(txtMapWidth.Text)
    M.MapHeight = CLng(txtMapHeight.Text)
    M.ViewLeft = CInt(txtDispLeft.Text)
    M.ViewTop = CInt(txtDispTop.Text)
    M.ViewWidth = CInt(txtDispWidth.Text)
    M.ViewHeight = CInt(txtDispHeight.Text)
    Prj.AddMap M
    UpdateMapList
    Set M = Nothing
    Exit Sub

MapCreateErr:
    MsgBox Err.Description
    Set M = Nothing
End Sub

Private Sub cmdSaveMap_Click()
    
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map to save"
        Exit Sub
    End If
    
    On Error Resume Next
    dlgMapFile.Flags = &H880E&
    dlgMapFile.InitDir = GetSetting("GameDev", "Directories", "MapPath", App.Path)
    dlgMapFile.ShowSave
    If Err.Number = 0 Then
        On Error GoTo SaveErr
        If GetRelativePath(Prj.ProjectPath, dlgMapFile.FileName) <> Prj.Maps(lstMaps.ListIndex).Path Then
            Prj.IsDirty = True ' Map file name changed
        End If
        Prj.Maps(lstMaps.ListIndex).Save GetRelativePath(Prj.ProjectPath, dlgMapFile.FileName)
        SaveSetting "GameDev", "Directories", "MapPath", Left$(dlgMapFile.FileName, Len(dlgMapFile.FileName) - Len(dlgMapFile.FileTitle) - 1)
    End If
    Exit Sub

SaveErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdUpdateFunc_Click()
    On Error GoTo UpdateFuncErr
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstSpecialFunctions.ListIndex < 0 Then
        MsgBox "Please select a special function to update first."
        Exit Sub
    End If
    With Prj.Maps(lstMaps.ListIndex)
        If Not .SpecialExists(lstSpecialFunctions.Text) Then
            MsgBox "Special function " & lstSpecialFunctions.Text & " no longer exists."
            Exit Sub
        End If
        .IsDirty = True
        With .Specials(lstSpecialFunctions.Text)
            .FuncType = cboFunction.ItemData(cboFunction.ListIndex)
            .Flags = IIf(chkFuncUp.Value = vbChecked, InteractionFlags.INTFL_ACTONUP, 0)
            If chkFuncButton.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_ACTONBUTTON
            If chkFuncDown.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_ACTONDOWN
            If chkRemoveFunc.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_REMOVEALWAYS
            If chkFuncInitial.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_INITIALTOUCH
            If chkFuncUseInvRemove.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_FUNCREMOVEINV
            If cboFuncUseInv.ListIndex >= 0 Then .InvItem = cboFuncUseInv.ListIndex
            If chkActOnStartup.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_ACTONSTARTUP
            .InvUseCount = Val(txtFuncUseInvCount.Text)
            .Value = ""
            .SpriteName = ""
            Select Case .FuncType
            Case SPECIAL_MESSAGE
                .Value = txtFuncMessage.Text
            Case SPECIAL_SWITCHSPRITE
                If lstSwitchSprite.ListIndex < 0 Then
                    MsgBox "Please select a sprite first"
                    Exit Sub
                End If
                .SpriteName = lstSwitchSprite.Text
                If chkSSSwapControl.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_SWAPCONTROL
                If chkSSNewInstance.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_NEWINSTANCE
                If chkSSDeleteOld.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_DELETEOLD
                If chkSSSameLocation.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_OLDLOCATION
            Case SPECIAL_SWITCHMAP
                If lstSwitchMap.ListIndex < 0 Then
                    MsgBox "Please select a map to switch to first"
                    Exit Sub
                End If
                .Value = lstSwitchMap.Text
                If cboSMSprite.ListIndex > 0 Then
                    .SpriteName = cboSMSprite.Text
                Else
                    .SpriteName = ""
                End If
                If chkSMPosOverride.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_OVERRIDEPOSITION
                .DestX = Val(txtSMX.Text)
                .DestY = Val(txtSMY.Text)
            Case SPECIAL_TELEPORT
                .DestX = Val(txtTeleportX.Text)
                .DestY = Val(txtTeleportY.Text)
                If chkTeleportOffset.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_RELATIVETELEPORT
            Case SPECIAL_ALTERMAP
                .DestX = Val(txtCopyToX.Text)
                .DestY = Val(txtCopyToY.Text)
                If lstAlterMapFunc.ListIndex < 0 Then
                    MsgBox "Please select a function to copy tiles from"
                Else
                    .Value = lstAlterMapFunc.Text
                End If
            Case SPECIAL_CREATESPRITE
                .DestX = Val(txtCreateSpriteX.Text)
                .DestY = Val(txtCreateSpriteY.Text)
                If chkCreatePos.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_OVERRIDEPOSITION
                If optCreatePosPlrRel.Value Then .Flags = .Flags Or InteractionFlags.INTFL_RELATIVETOPLAYER
                If lstCreateSprite.ListIndex < 0 Then
                    MsgBox "Please select a sprite to create"
                Else
                    .SpriteName = lstCreateSprite.Text
                End If
                .Value = CStr(Val(txtCreateMaxCount.Text))
            Case SPECIAL_DELETESPRITE
                If lstDeleteSprite.ListIndex < 0 Then
                    MsgBox "Please select a sprite to delete"
                Else
                    .SpriteName = lstDeleteSprite.Text
                End If
                .Value = CStr(Val(txtDeleteCount.Text))
                If chkDeleteMany.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_DELETEMANY
            Case SPECIAL_EVENT
                .Flags = .Flags Or InteractionFlags.INTFL_RAISEEVENT
            End Select
        End With
    End With
    Exit Sub

UpdateFuncErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub LoadFunc(Optional bFuncChanging As Boolean = False)
    Dim bEnable As Boolean
    Dim Idx As Integer
    
    On Error GoTo LoadFuncErr
    
    If lstMaps.ListIndex < 0 Then fraSpecialFunctions.Enabled = False
    fraSpecialFunctions.Enabled = True
    fraFuncMessage.Visible = False
    fraSwitchSprite.Visible = False
    fraSwitchMap.Visible = False
    fraTeleport.Visible = False
    fraAlterMap.Visible = False
    fraCreateSprite.Visible = False
    fraDeleteSprite.Visible = False
    fraNotOnLayer.Visible = False
    
    With Prj.Maps(lstMaps.ListIndex)
        bEnable = Not (lstSpecialFunctions.ListIndex < 0)
        If bEnable Then bEnable = .SpecialExists(lstSpecialFunctions.Text)
        cmdUpdateFunc.Enabled = bEnable
        cboFunction.Enabled = bEnable
        chkFuncUp.Enabled = bEnable
        chkFuncButton.Enabled = bEnable
        chkFuncDown.Enabled = bEnable
        chkRemoveFunc.Enabled = bEnable
        chkFuncInitial.Enabled = bEnable
        chkFuncUseInventory.Enabled = bEnable
        If bEnable = False Then chkFuncUseInventory.Value = vbUnchecked: chkFuncUseInventory_Click
        cmdDeleteFunc.Enabled = bEnable
        If Not bEnable Then Exit Sub
        With .Specials(lstSpecialFunctions.Text)
            If Not bFuncChanging Then
                For Idx = 0 To cboFunction.ListCount - 1
                    If cboFunction.ItemData(Idx) = .FuncType Then
                        cboFunction.ListIndex = Idx
                        Exit For
                    End If
                Next
            End If
            If PlayerSpriteDef Is Nothing Then
                fraNotOnLayer.Visible = True
            Else
                If Not (Prj.Maps(lstMaps.ListIndex).MapLayer(.LayerIndex) Is PlayerSpriteDef.rLayer) Then
                    fraNotOnLayer.Visible = True
                End If
            End If
            If cboFunction.ListIndex >= 0 Then
                Select Case cboFunction.ItemData(cboFunction.ListIndex)
                Case SpecialFuncs.SPECIAL_MESSAGE
                    fraFuncMessage.Visible = True
                Case SpecialFuncs.SPECIAL_SWITCHSPRITE
                    fraSwitchSprite.Visible = True
                Case SpecialFuncs.SPECIAL_SWITCHMAP
                    fraSwitchMap.Visible = True
                Case SpecialFuncs.SPECIAL_TELEPORT
                    fraTeleport.Visible = True
                Case SpecialFuncs.SPECIAL_ALTERMAP
                    fraAlterMap.Visible = True
                Case SpecialFuncs.SPECIAL_CREATESPRITE
                    fraCreateSprite.Visible = True
                Case SpecialFuncs.SPECIAL_DELETESPRITE
                    fraDeleteSprite.Visible = True
                End Select
            End If
            If Not bFuncChanging Then
                chkFuncUp.Value = IIf(.Flags And InteractionFlags.INTFL_ACTONUP, vbChecked, vbUnchecked)
                chkFuncButton.Value = IIf(.Flags And InteractionFlags.INTFL_ACTONBUTTON, vbChecked, vbUnchecked)
                chkFuncDown.Value = IIf(.Flags And InteractionFlags.INTFL_ACTONDOWN, vbChecked, vbUnchecked)
                chkFuncInitial.Value = IIf(.Flags And InteractionFlags.INTFL_INITIALTOUCH, vbChecked, vbUnchecked)
                chkRemoveFunc.Value = IIf(.Flags And InteractionFlags.INTFL_REMOVEALWAYS, vbChecked, vbUnchecked)
                chkFuncUseInventory.Value = IIf(.InvUseCount > 0, vbChecked, vbUnchecked)
                cboFuncUseInv.ListIndex = IIf(.InvUseCount > 0, .InvItem, -1)
                chkActOnStartup.Value = IIf(.Flags And InteractionFlags.INTFL_ACTONSTARTUP, vbChecked, vbUnchecked)
                txtFuncUseInvCount.Text = CStr(.InvUseCount)
                chkFuncUseInvRemove.Value = IIf(.Flags And InteractionFlags.INTFL_FUNCREMOVEINV, vbChecked, vbUnchecked)
            End If
            txtFuncMessage.Text = .Value
            lstSwitchSprite.Clear
            For Idx = 0 To Prj.Maps(lstMaps.ListIndex).SpriteDefCount - 1
                lstSwitchSprite.AddItem Prj.Maps(lstMaps.ListIndex).SpriteDefs(Idx).Name
                If Prj.Maps(lstMaps.ListIndex).SpriteDefs(Idx).Name = .SpriteName Then
                    lstSwitchSprite.ListIndex = Idx
                End If
            Next
            chkSSSwapControl.Value = IIf(.Flags And InteractionFlags.INTFL_SWAPCONTROL, vbChecked, vbUnchecked)
            chkSSNewInstance.Value = IIf(.Flags And InteractionFlags.INTFL_NEWINSTANCE, vbChecked, vbUnchecked)
            chkSSDeleteOld.Value = IIf(.Flags And InteractionFlags.INTFL_DELETEOLD, vbChecked, vbUnchecked)
            chkSSSameLocation.Value = IIf(.Flags And InteractionFlags.INTFL_OLDLOCATION, vbChecked, vbUnchecked)
            lstSwitchMap.Clear
            For Idx = 0 To Prj.MapCount - 1
                lstSwitchMap.AddItem Prj.Maps(Idx).Name
                If Prj.Maps(Idx).Name = .Value Then
                    lstSwitchMap.ListIndex = Idx
                End If
            Next
            FillSwitchMapSprites
            If lstSwitchMap.ListIndex >= 0 Then
                If .SpriteName = "" Then
                    cboSMSprite.ListIndex = 0
                Else
                    For Idx = 0 To cboSMSprite.ListCount - 1
                        If .SpriteName = cboSMSprite.List(Idx) Then
                            cboSMSprite.ListIndex = Idx
                            Exit For
                        End If
                    Next
                End If
            End If
            lstAlterMapFunc.Clear
            With Prj.Maps(lstMaps.ListIndex)
                For Idx = 0 To .SpecialCount - 1
                    lstAlterMapFunc.AddItem .Specials(Idx).Name
                    If .Specials(Idx).Name = .Specials(lstSpecialFunctions.Text).Value Then
                        lstAlterMapFunc.ListIndex = Idx
                    End If
                Next
            End With
            lstCreateSprite.Clear
            lstDeleteSprite.Clear
            With Prj.Maps(lstMaps.ListIndex)
                For Idx = 0 To .SpriteDefCount - 1
                    lstCreateSprite.AddItem .SpriteDefs(Idx).Name
                    lstDeleteSprite.AddItem .SpriteDefs(Idx).Name
                    If .SpriteDefs(Idx).Name = .Specials(lstSpecialFunctions.Text).SpriteName Then
                        lstCreateSprite.ListIndex = Idx
                        lstDeleteSprite.ListIndex = Idx
                    End If
                Next
            End With
            chkSMPosOverride.Value = IIf(.Flags And InteractionFlags.INTFL_OVERRIDEPOSITION, vbChecked, vbUnchecked)
            txtSMX.Text = CStr(.DestX)
            txtSMY.Text = CStr(.DestY)
            txtTeleportX.Text = CStr(.DestX)
            txtTeleportY.Text = CStr(.DestY)
            txtCopyToX.Text = CStr(.DestX)
            txtCopyToY.Text = CStr(.DestY)
            txtCreateSpriteX.Text = CStr(.DestX)
            txtCreateSpriteY.Text = CStr(.DestY)
            txtCreateMaxCount.Text = .Value
            txtDeleteCount.Text = .Value
            chkDeleteMany.Value = IIf(.Flags And InteractionFlags.INTFL_DELETEMANY, vbChecked, vbUnchecked)
            cboPickCoord.ListIndex = -1
            chkTeleportOffset.Value = IIf(.Flags And InteractionFlags.INTFL_RELATIVETELEPORT, vbChecked, vbUnchecked)
            chkCreatePos.Value = IIf(.Flags And InteractionFlags.INTFL_OVERRIDEPOSITION, vbChecked, vbUnchecked)
            chkCreatePos_Click
            If .Flags And InteractionFlags.INTFL_RELATIVETOPLAYER Then optCreatePosPlrRel.Value = True Else optCreatePosAbsolute.Value = True
            DrawTeleportPreview
        End With
    End With
    Exit Sub

LoadFuncErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub FillSwitchMapSprites()
    Dim Idx As Integer

    On Error GoTo FillSwitchMapErr
    If lstSwitchMap.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstSwitchMap.ListIndex)
        cboSMSprite.Clear
        If lstSwitchMap.ListIndex >= 0 Then
            cboSMSprite.AddItem "(default)"
            For Idx = 0 To .SpriteDefCount - 1
                cboSMSprite.AddItem .SpriteDefs(Idx).Name
            Next
        End If
    End With
    Exit Sub
    
FillSwitchMapErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub DrawTeleportPreview()
    Dim TmpImg As StdPicture
    Dim TPX As Long, TPY As Long
    Dim TLX As Long, TLY As Long
    Dim TLTX As Long, TLTY As Long
    Dim XDraw As Long, YDraw As Long
    
    If PlayerSpriteDef Is Nothing Then Exit Sub
    If chkTeleportOffset.Value = vbChecked Then
        picTeleport.Cls
        picTeleport.Print "Preview unavailable"
        Exit Sub
    End If
    With PlayerSpriteDef.rLayer
        If .TSDef.IsLoaded Then
            Set TmpImg = .TSDef.Image
        Else
            Set TmpImg = LoadPicture(.TSDef.ImagePath)
        End If
        TPX = Val(txtTeleportX.Text)
        TLX = (TPX - picTeleport.ScaleWidth \ 2)
        If TLX < 0 Then TLX = 0
        If (TLX + picTeleport.ScaleWidth) \ .TSDef.TileWidth > .Columns - 1 Then TLX = .Columns * .TSDef.TileWidth - 1 - picTeleport.ScaleWidth
        TPY = Val(txtTeleportY.Text)
        TLY = (TPY - picTeleport.ScaleHeight \ 2)
        If TLY < 0 Then TLY = 0
        If (TLY + picTeleport.ScaleHeight) \ .TSDef.TileHeight > .Rows - 1 Then TLY = .Rows * .TSDef.TileHeight - 1 - picTeleport.ScaleHeight
        For YDraw = TLY \ .TSDef.TileHeight To (TLY + picTeleport.ScaleHeight) \ .TSDef.TileHeight
            For XDraw = TLX \ .TSDef.TileWidth To (TLX + picTeleport.ScaleWidth) \ .TSDef.TileWidth
                picTeleport.PaintPicture ExtractTilesetTile(.TSDef, TmpImg, .Data.TileValue(XDraw, YDraw)), XDraw * .TSDef.TileWidth - TLX, YDraw * .TSDef.TileHeight - TLY
            Next
        Next
        picTeleport.Line (TPX - 1 - TLX, TPY - 1 - TLY)-(TPX + 1 - TLX, TPY + 1 - TLY), vbWhite, B
        picTeleport.PSet (TPX - TLX, TPY - TLY), vbRed
    End With
End Sub

Private Sub cmdUpdateLayer_Click()
    Dim Index As Integer
    
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map whose layer you wish to update before selecting this command", vbExclamation
        Exit Sub
    End If
    If lstLayers.ListIndex < 0 Then
        MsgBox "Please select a layer to update before selecting this command", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo UpdateLayerErr
    With Prj.Maps(lstMaps.ListIndex).MapLayer(lstLayers.ListIndex)
        .UpdateLayer txtLayerName.Text, 0, 0, Prj.TileSetDef(cboLayerTileset.ListIndex), CSng(txtYRate.Text), CSng(txtYRate.Text), IIf(chkTransparent.Value = vbChecked, True, False)
    End With
    Exit Sub

UpdateLayerErr:
    MsgBox "Error updating layer: " & Err.Description, vbExclamation
End Sub

Private Sub cmdUpdateMap_Click()
    Dim li As Integer

    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map to update before selecting this command"
        Exit Sub
    End If
    
    On Error GoTo UpdateMapErr
    With Prj.Maps(lstMaps.ListIndex)
        .Name = txtMapName.Text
        .MapWidth = CLng(txtMapWidth.Text)
        .MapHeight = CLng(txtMapHeight.Text)
        .ViewLeft = CInt(txtDispLeft.Text)
        .ViewTop = CInt(txtDispTop.Text)
        .ViewWidth = CInt(txtDispWidth.Text)
        .ViewHeight = CInt(txtDispHeight.Text)
        For li = 0 To .LayerCount - 1
            .MapLayer(li).UpdateLayer .MapLayer(li).Name, .MapWidth, .MapHeight, .MapLayer(li).TSDef, .MapLayer(li).XScrollRate, .MapLayer(li).YScrollRate, .MapLayer(li).Transparent
        Next
    End With
        
    Exit Sub

UpdateMapErr:
    MsgBox "Error updating map: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "MapEdit", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    Select Case Val(GetSetting("GameDev", "Options", "ScreenDepth", "16"))
    Case 24
        opt24Bit.Value = True
    Case 32
        opt32Bit.Value = True
    Case Else
        opt16Bit.Value = True
    End Select
    FillTilesets
    UpdateMapList
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub FillTilesets()
    Dim I As Integer
    
    cboLayerTileset.Clear
    
    For I = 0 To Prj.TileSetDefCount - 1
        cboLayerTileset.AddItem Prj.TileSetDef(I).Name
    Next I
    
End Sub

Sub UpdateMapList()
    Dim Index As Integer
    
    lstMaps.Clear
    For Index = 0 To Prj.MapCount - 1
        lstMaps.AddItem Prj.Maps(Index).Name
    Next
    
End Sub

Sub UpdateLayerList()
    Dim Index As Integer
    Dim M As Map
    
    lstLayers.Clear
    If lstMaps.ListIndex < 0 Then Exit Sub
    Set M = Prj.Maps(lstMaps.ListIndex)
    For Index = 0 To M.LayerCount - 1
        lstLayers.AddItem M.MapLayer(Index).Name
    Next Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "MapEdit", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub lstLayers_Click()
    Dim LYR As Layer
    Dim Index As Integer
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    If lstLayers.ListIndex < 0 Then Exit Sub
    Set LYR = Prj.Maps(lstMaps.ListIndex).MapLayer(lstLayers.ListIndex)
    For Index = 0 To Prj.TileSetDefCount - 1
        If Prj.TileSetDef(Index) Is LYR.TSDef Then
            cboLayerTileset.ListIndex = Index
        End If
    Next
    txtLayerName.Text = LYR.Name
    txtXRate.Text = CStr(LYR.XScrollRate)
    txtYRate.Text = CStr(LYR.YScrollRate)
    chkTransparent.Value = IIf(LYR.Transparent, vbChecked, vbUnchecked)
End Sub

Private Sub lstMaps_Click()
    If lstMaps.ListIndex < 0 Then Exit Sub
    UpdateLayerList
    With Prj.Maps(lstMaps.ListIndex)
        txtMapName.Text = .Name
        lblMapPath.Caption = .Path
        txtMapWidth.Text = CStr(.MapWidth)
        txtMapHeight.Text = CStr(.MapHeight)
        txtDispLeft.Text = CStr(.ViewLeft)
        txtDispTop.Text = CStr(.ViewTop)
        txtDispWidth.Text = CStr(.ViewWidth)
        txtDispHeight.Text = CStr(.ViewHeight)
    End With
    
    On Error Resume Next
    LoadPlayerSprite
    FillCategories
    FillInventory
    LoadReplaceTile
    FillTouchMedia
    LoadInteraction
    If tabMaps.SelectedItem.Key = "SpecialFunctions" Then FillSpecialFunctions
    
End Sub

Private Sub cmdDeleteMI_Click()
    If lstMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        .RemoveInteraction nCurInt
        If nCurInt > .InteractCount - 1 Then
            nCurInt = .InteractCount - 1
        End If
        LoadInteraction
        .IsDirty = True
    End With
End Sub

Private Sub cmdFirstMI_Click()
    StoreInteraction
    nCurInt = 0
    LoadInteraction
End Sub

Private Sub cmdLastMI_Click()
    If lstMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        StoreInteraction
        nCurInt = .InteractCount - 1
        LoadInteraction
    End With
End Sub

Private Sub cmdNewMI_Click()
    Dim NewInt As New Interaction
    
    If lstMaps.ListIndex < 0 Then
        MsgBox "Please select a map before selecting this command", vbExclamation
        Exit Sub
    End If
    With Prj.Maps(lstMaps.ListIndex)
        nCurInt = .InteractCount
        .AddInteraction NewInt
        StoreInteraction
        Set NewInt = Nothing
        LoadInteraction
    End With
End Sub

Private Sub cmdNextMI_Click()
    StoreInteraction
    nCurInt = nCurInt + 1
    LoadInteraction
End Sub

Private Sub cmdPrevMI_Click()
    StoreInteraction
    nCurInt = nCurInt - 1
    LoadInteraction
End Sub

Private Sub cmdUpdateMI_Click()
    StoreInteraction
End Sub

Public Sub LoadInteraction()
    Dim Idx As Integer

    On Error GoTo LoadIntErr
    
    If lstMaps.ListIndex < 0 Then
        fraPlayerInteraction.Caption = "No Map Selected"
        cmdFirstMI.Enabled = False
        cmdPrevMI.Enabled = False
        cmdNextMI.Enabled = False
        cmdLastMI.Enabled = False
        cmdDeleteMI.Enabled = False
        cmdUpdateMI.Enabled = False
        Exit Sub
    End If
    
    With Prj.Maps(lstMaps.ListIndex)
    
        If .InteractCount <= 0 Then
            fraPlayerInteraction.Caption = "No Interactions Defined"
            cmdFirstMI.Enabled = False
            cmdPrevMI.Enabled = False
            cmdNextMI.Enabled = False
            cmdLastMI.Enabled = False
            cmdDeleteMI.Enabled = False
            cmdUpdateMI.Enabled = False
            Exit Sub
        End If
        If nCurInt >= .InteractCount Then nCurInt = .InteractCount - 1
        fraPlayerInteraction.Caption = "Map Interaction " & CStr(nCurInt) + 1 & " of " & .InteractCount
        With .Interactions(nCurInt)
            For Idx = 0 To cboTouchCategory.ListCount - 1
                If .TouchCategory.Name = cboTouchCategory.List(Idx) Then
                    cboTouchCategory.ListIndex = Idx
                    Exit For
                End If
            Next
            If .Flags And InteractionFlags.INTFL_INITIALTOUCH Then
                optFirstTouch.Value = True
            Else
                optContinuousTouch.Value = True
            End If
            If .Flags And InteractionFlags.INTFL_REMOVEALWAYS Then
                optAlwaysRemove.Value = True
            ElseIf .Flags And InteractionFlags.INTFL_REMOVEIFACT Then
                optRemoveIfUsed.Value = True
            Else
                optDontRemove.Value = True
            End If
            cboReaction.ListIndex = .Reaction
            cboRelaventInventory.ListIndex = .InvItem
            txtReplaceTile.Text = CStr(.ReplaceTile)
            updReplaceTile.Value = .ReplaceTile
            cboTouchMedia.ListIndex = 0
            For Idx = 1 To cboTouchMedia.ListCount - 1
                If cboTouchMedia.List(Idx) = .Media Then
                    cboTouchMedia.ListIndex = Idx
                    Exit For
                End If
            Next
            chkRaiseEvent.Value = IIf(.Flags And InteractionFlags.INTFL_RAISEEVENT, vbChecked, vbUnchecked)
        End With
        If nCurInt <= 0 Then cmdPrevMI.Enabled = False Else cmdPrevMI.Enabled = True
        If nCurInt >= .InteractCount - 1 Then cmdNextMI.Enabled = False Else cmdNextMI.Enabled = True
    End With
    
    cmdFirstMI.Enabled = True
    cmdLastMI.Enabled = True
    cmdDeleteMI.Enabled = True
    cmdUpdateMI.Enabled = True
    Exit Sub
    
LoadIntErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub StoreInteraction()
    On Error GoTo StoreIErr
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        If PlayerSpriteDef Is Nothing Then Exit Sub
        If .InteractCount <= 0 Then Exit Sub
        
        .IsDirty = True
        With .Interactions(nCurInt)
            Set .TouchCategory = Prj.Groups(cboTouchCategory.Text, PlayerSpriteDef.rLayer.TSDef.Name)
            .Flags = IIf(optFirstTouch.Value, InteractionFlags.INTFL_INITIALTOUCH, 0)
            .Flags = .Flags Or IIf(optRemoveIfUsed.Value, InteractionFlags.INTFL_REMOVEIFACT, IIf(optAlwaysRemove.Value, InteractionFlags.INTFL_REMOVEALWAYS, 0))
            If chkRaiseEvent.Value = vbChecked Then .Flags = .Flags Or InteractionFlags.INTFL_RAISEEVENT
            .InvItem = cboRelaventInventory.ListIndex
            .Reaction = cboReaction.ListIndex
            .ReplaceTile = updReplaceTile.Value
            .Media = IIf(cboTouchMedia.ListIndex > 0, cboTouchMedia.Text, "")
        End With
    End With
    Exit Sub

StoreIErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub lstSpecialFunctions_Click()
    LoadFunc False
End Sub

Private Sub lstSwitchMap_Click()
    FillSwitchMapSprites
End Sub

Private Sub tabFunctions_Click()
    If tabFunctions.SelectedItem.Index = 1 Then
        fraFuncAct.Visible = True
        fraFuncDef.Visible = False
    Else
        fraFuncAct.Visible = False
        fraFuncDef.Visible = True
    End If
End Sub

Private Sub tabMaps_Click()
    On Error GoTo TabSwitchErr

    fraLayers.Visible = False
    fraPlayerInteraction.Visible = False
    fraSpecialFunctions.Visible = False
    fraPlayerSprite.Visible = False
    
    If tabMaps.SelectedItem.Key = "Layers" Then
        fraLayers.Visible = True
    ElseIf tabMaps.SelectedItem.Key = "PlayerInteraction" Then
        fraPlayerSprite.Visible = True
        fraPlayerInteraction.Visible = True
        LoadPlayerSprite
        FillCategories
        FillInventory
        LoadReplaceTile
        FillTouchMedia
        LoadInteraction
    Else
        fraSpecialFunctions.Visible = True
        FillSpecialFunctions
        FillFuncInvItems
        chkFuncUseInventory_Click
    End If
    Exit Sub

TabSwitchErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub FillTouchMedia()
    Dim Idx As Integer
    
    cboTouchMedia.Clear
    cboTouchMedia.AddItem "(None)"
    For Idx = 0 To Prj.MediaMgr.MediaClipCount - 1
        cboTouchMedia.AddItem Prj.MediaMgr.Clip(Idx).Name
    Next
End Sub

Public Sub FillFuncInvItems()
    Dim Idx As Integer
    
    On Error GoTo FillInvErr
    cboFuncUseInv.Clear
    For Idx = 0 To Prj.GamePlayer.InventoryCount - 1
        cboFuncUseInv.AddItem Prj.GamePlayer.InventoryItemName(Idx)
    Next
    Exit Sub

FillInvErr:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub LoadPlayerSprite()
    Dim Idx As Integer
    
    If lstMaps.ListIndex < 0 Then Exit Sub
    With Prj.Maps(lstMaps.ListIndex)
        cboSprites.Clear
        For Idx = 0 To .SpriteDefCount - 1
            If .SpriteDefs(Idx).Flags And eDefFlagBits.FLAG_INSTANCE Then
                cboSprites.AddItem .SpriteDefs(Idx).Name
                If .SpriteDefs(Idx).Name = .PlayerSpriteName Then
                    cboSprites.ListIndex = cboSprites.NewIndex
                End If
            End If
        Next
    End With
End Sub

Public Sub LoadReplaceTile()
    On Error Resume Next
    If PlayerSpriteDef Is Nothing Then
        Set imgReplaceTile.Picture = LoadPicture()
        Exit Sub
    End If
    With PlayerSpriteDef.rLayer.TSDef
        If .Image Is Nothing Then
            .Load
        End If
        updReplaceTile.Max = (ScaleX(.Image.Width, vbHimetric, vbPixels) / .TileWidth) * _
            (ScaleY(.Image.Height, vbHimetric, vbPixels) / .TileHeight) - 1
    End With
    Set imgReplaceTile.Picture = ExtractReplaceTile(updReplaceTile.Value)
    If Err.Number Then
        MsgBox Err.Description, vbExclamation
    End If
End Sub

Public Sub FillInventory()
    Dim Idx As Integer
    
    cboRelaventInventory.Clear
    For Idx = 0 To Prj.GamePlayer.InventoryCount - 1
        cboRelaventInventory.AddItem Prj.GamePlayer.InventoryItemName(Idx)
    Next
End Sub

Public Sub FillCategories()
    Dim Idx As Integer
    
    cboTouchCategory.Clear
    If Not PlayerSpriteDef Is Nothing Then
        With PlayerSpriteDef.rLayer.TSDef
            For Idx = 0 To Prj.GroupByTilesetCount(.Name) - 1
                cboTouchCategory.AddItem Prj.TilesetGroupByIndex(.Name, Idx).Name
            Next
        End With
    End If

End Sub

Public Sub FillSpecialFunctions()
    Dim Idx As Integer
    Dim PSD As SpriteDef
    
    If PlayerSpriteDef Is Nothing Then
        MsgBox "A map must be selected and the player sprite specified to use special functions."
        fraSpecialFunctions.Enabled = False
        Exit Sub
    End If
    fraSpecialFunctions.Enabled = True
        
    lstSpecialFunctions.Clear
    Set PSD = PlayerSpriteDef
    With PSD.rLayer.pMap
        For Idx = 0 To .SpecialCount - 1
            'If PSD.rLayer Is .MapLayer(.Specials(Idx).LayerIndex) Then
                lstSpecialFunctions.AddItem .Specials(Idx).Name
            'End If
        Next
    End With
    Set PSD = Nothing
End Sub

Public Function PlayerSpriteDef() As SpriteDef
    Dim Idx As Integer
    
    If lstMaps.ListIndex < 0 Then Exit Function
    With Prj.Maps(lstMaps.ListIndex)
        For Idx = 0 To .SpriteDefCount - 1
            If .SpriteDefs(Idx).Name = cboSprites.Text Then
                Set PlayerSpriteDef = .SpriteDefs(Idx)
            End If
        Next
    End With
End Function

Private Function ExtractReplaceTile(ByVal Index As Integer) As StdPicture
    Dim TSCols As Integer
    Dim TSRows As Integer
    
    On Error GoTo ExtractErr
    
    If PlayerSpriteDef Is Nothing Then
        Set ExtractReplaceTile = LoadPicture()
        Exit Function
    End If
    
    With PlayerSpriteDef.rLayer.TSDef
        If .Image Is Nothing Then
            .Load
        End If
        If .Image Is Nothing Then Exit Function
        TSCols = ScaleX(.Image.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(.Image.Height, vbHimetric, vbPixels) \ .TileHeight
        If Index < TSRows * TSCols Then
            Set ExtractReplaceTile = ExtractTile(.Image, .TileWidth * (Index Mod TSCols), .TileHeight * (Index \ TSCols), .TileWidth, .TileHeight)
        Else
            MsgBox "Tile index out of bounds", vbExclamation, "ExtractReplaceTile"
        End If
    End With
    Exit Function

ExtractErr:
    MsgBox Err.Description, vbExclamation
End Function

Private Function ExtractTilesetTile(TS As TileSetDef, Img As StdPicture, ByVal Index As Integer) As StdPicture
    Dim TSCols As Integer
    Dim TSRows As Integer
    
    On Error GoTo ExtractErr

    With TS
        TSCols = ScaleX(Img.Width, vbHimetric, vbPixels) \ .TileWidth
        TSRows = ScaleY(Img.Height, vbHimetric, vbPixels) \ .TileHeight
        If Index < TSRows * TSCols Then
            Set ExtractTilesetTile = ExtractTile(Img, .TileWidth * (Index Mod TSCols), .TileHeight * (Index \ TSCols), .TileWidth, .TileHeight)
        Else
            MsgBox "Tile index out of bounds", vbExclamation, "ExtractTilesetTile"
        End If
    End With
    Exit Function

ExtractErr:
    MsgBox Err.Description, vbExclamation
End Function

Private Sub txtReplaceTile_Change()
    On Error Resume Next
    updReplaceTile.Value = CInt(txtReplaceTile.Text)
End Sub

Private Sub txtTeleportX_Change()
    DrawTeleportPreview
End Sub

Private Sub txtTeleportY_Change()
    DrawTeleportPreview
End Sub

Private Sub updReplaceTile_Change()
    LoadReplaceTile
End Sub
