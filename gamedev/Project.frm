VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProject 
   Caption         =   "Scrolling Game Development Kit"
   ClientHeight    =   3855
   ClientLeft      =   165
   ClientTop       =   720
   ClientWidth     =   7260
   HelpContextID   =   112
   Icon            =   "Project.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tbrProject 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "Create new project"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open an existing project"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save the current project"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Description     =   "Export"
            Object.ToolTipText     =   "Export the current project (GDP + MAPs) as XML"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Import"
            Description     =   "Import"
            Object.ToolTipText     =   "Import a GameDev project from XML"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Shortcut"
            Description     =   "Make shortcut"
            Object.ToolTipText     =   "Make a shortcut to play the project"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            Description     =   "Play"
            Object.ToolTipText     =   "Play the current project"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tilesets"
            Description     =   "Tilesets"
            Object.ToolTipText     =   "Edit tilesets"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Maps"
            Description     =   "Maps"
            Object.ToolTipText     =   "Edit maps"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileMatch"
            Description     =   "Tile matching"
            Object.ToolTipText     =   "Tile matching"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileAnim"
            Description     =   "Tile animation"
            Object.ToolTipText     =   "Tile animation"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Categories"
            Description     =   "Tile categories and solidity"
            Object.ToolTipText     =   "Tile categories and solidity"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sprites"
            Description     =   "Sprites and paths"
            Object.ToolTipText     =   "Sprites and paths"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CollDef"
            Description     =   "Collisions"
            Object.ToolTipText     =   "Define collisions"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Player"
            Description     =   "Player settings"
            Object.ToolTipText     =   "Player settings and inventory"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Media"
            Description     =   "Media clips"
            Object.ToolTipText     =   "Manage media clips"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Controls"
            Description     =   "Controller configuration"
            Object.ToolTipText     =   "Controller configuration"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Description     =   "Options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Description     =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   23
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlXML 
      Left            =   2280
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".xml"
      Filter          =   "XML files (*.xml)|*.xml|All Files (*.*)|*.*"
      FilterIndex     =   1
   End
   Begin MSComctlLib.TreeView tvwProject 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5212
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgProjectFile 
      Left            =   1680
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".gdp"
      DialogTitle     =   "Specify project file name"
      Filter          =   "All Files (*.*)|*.*|GameDev Projects (*.gdp)|*.gdp"
      FilterIndex     =   2
      Flags           =   34822
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8388736
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":0E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":0F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1030
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1142
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1254
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1366
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1478
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":169C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":17AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":19D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":1F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":2046
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":2158
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":226A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":237C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":248E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":25A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Project.frx":26B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgMouse 
      Height          =   255
      Left            =   0
      Picture         =   "Project.frx":27C4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export XML"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import XML"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayGame 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortcut 
         Caption         =   "&Make Shortcut"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuTilesets 
         Caption         =   "&Tilesets"
      End
      Begin VB.Menu mnuMaps 
         Caption         =   "&Maps"
      End
      Begin VB.Menu mnuMatching 
         Caption         =   "Ti&le Matching"
      End
      Begin VB.Menu mnuAnimation 
         Caption         =   "Tile &Animation"
      End
      Begin VB.Menu mnuCategories 
         Caption         =   "Tile &Categories"
      End
      Begin VB.Menu mnuSprites 
         Caption         =   "&Sprites and Paths"
      End
      Begin VB.Menu mnuCollisions 
         Caption         =   "Collision &Definitions"
      End
      Begin VB.Menu mnuPlayer 
         Caption         =   "&Player Settings"
      End
      Begin VB.Menu mnuMedia 
         Caption         =   "M&edia Clips"
      End
      Begin VB.Menu mnuCtrlConfig 
         Caption         =   "Co&ntroller Configuration"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
      End
      Begin VB.Menu mnuHelpTutorial 
         Caption         =   "Tuto&rial"
      End
      Begin VB.Menu mnuQuickTut 
         Caption         =   "&Quick Start Tutorial"
      End
      Begin VB.Menu mnuHelpTech 
         Caption         =   "&Technical Support Info"
      End
      Begin VB.Menu mnuSeparatorH1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuGoto 
         Caption         =   "&Go to Item"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Tree"
      End
   End
End
Attribute VB_Name = "frmProject"
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
' File: Project.frm - Main Window; Project Tree Dialog
'
'======================================================================

Option Explicit

Private Const HH_HELP_TOC = &H1
Private Const HH_HELP_INDEX = &H2
Private Const HH_HELP_CONTEXT = &HF       ' Display mapped numeric value in
                                          ' dwData.
                                          ' WinHelp's HELP_WM_HELP.

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
                (ByVal hwndCaller As Long, ByVal pszFile As String, _
                 ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Sub Form_Load()
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "Project", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6)), CLng(Mid$(WndPos, 15, 6)), CLng(Right$(WndPos, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    Exit Sub

LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not UnloadProject
End Sub

' Returns True if Project can be safely unloaded
Private Function UnloadProject() As Boolean
    Dim Dirty As String
    Dim Index As Integer

QueryAgain:
    Dirty = ""
    If Prj.IsDirty Then Dirty = Dirty & Prj.ProjectPath & vbCrLf
    For Index = 0 To Prj.MapCount - 1
        If Prj.Maps(Index).IsDirty Then Dirty = Dirty & Prj.Maps(Index).Path & " (" & Prj.Maps(Index).Name & ")" & vbCrLf
    Next
    
    For Index = 0 To Prj.TileSetDefCount - 1
        If Prj.TileSetDef(Index).IsDirty Then Dirty = Dirty & Prj.TileSetDef(Index).ImagePath & " (" & Prj.TileSetDef(Index).Name & ")" & vbCrLf
    Next
    
    If Len(Dirty) Then
        Select Case MsgBox("The following files have changed, save changes?" & vbCrLf & Dirty, vbYesNoCancel)
        Case vbYes
            Prj.Save Prj.ProjectPath
            GoTo QueryAgain
        Case vbNo
            UnloadProject = True
        Case vbCancel
            UnloadProject = False
        End Select
    Else
        UnloadProject = True
    End If
End Function

Private Sub Form_Resize()
    On Error Resume Next
    tvwProject.Move tvwProject.Left, tbrProject.Height, Me.ScaleWidth - tvwProject.Left * 2, Me.ScaleHeight - tbrProject.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim F As Form
    For Each F In Forms
        If Not (F Is Me) Then Unload F
    Next
    SaveSetting "GameDev", "Windows", "Project", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000") & "," & Format$(Me.Width, " 00000;-00000") & "," & Format$(Me.Height, " 00000;-00000")
End Sub

Public Sub LoadTree()
    Dim nodeProject As Node
    Dim nodeProjElem As Node
    Dim nodeMap As Node
    Dim nodeTileset As Node
    Dim nodeMatch As Node
    Dim nodeAnim As Node
    Dim nodeLayers As Node
    Dim nodeLayer As Node
    Dim nodeCategories As Node
    Dim nodeSprites As Node
    Dim nodePaths As Node
    Dim nodePath As Node
    Dim nodeSolidDefs As Node
    Dim nodeTemplates As Node
    Dim nodeTemplate As Node
    Dim nodeMedia As Node
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim L As Integer

    tvwProject.Nodes.Clear
    Set tvwProject.ImageList = imlIcons
    Set nodeProject = tvwProject.Nodes.Add(, , "ProjectRoot", IIf(Prj.ProjectPath = "", "New Project", Prj.ProjectPath), 1)
    nodeProject.ExpandedImage = 2
    Set nodeProjElem = tvwProject.Nodes.Add(nodeProject, tvwChild, "MapRoot", "Maps", 1)
    nodeProjElem.EnsureVisible
    nodeProjElem.ExpandedImage = 2
    For I = 0 To Prj.MapCount - 1
        Set nodeMap = tvwProject.Nodes.Add(nodeProjElem, tvwChild, , Prj.Maps(I).Name, 3)
        nodeMap.Tag = "MP" & CStr(I)
        With Prj.Maps(I)
            If .LayerCount > 0 Then
                Set nodeLayers = tvwProject.Nodes.Add(nodeMap, tvwChild, , "Layers", 1)
                nodeLayers.ExpandedImage = 2
                For J = 0 To .LayerCount - 1
                    Set nodeLayer = tvwProject.Nodes.Add(nodeLayers, tvwChild, , .MapLayer(J).Name, 9)
                    nodeLayer.Tag = "LY" & CStr(I) & "," & CStr(J)
                    Set nodeAnim = Nothing
                    For K = 0 To Prj.AnimDefCount - 1
                        If Prj.AnimDefs(K).MapName = .Name And Prj.AnimDefs(K).LayerName = .MapLayer(J).Name Then
                            If nodeAnim Is Nothing Then
                                Set nodeAnim = tvwProject.Nodes.Add(nodeLayer, tvwChild, , "Tile Animations", 1)
                                nodeAnim.ExpandedImage = 2
                            End If
                            tvwProject.Nodes.Add(nodeAnim, tvwChild, , Prj.AnimDefs(K).Name, 7).Tag = "AD" & CStr(K)
                        End If
                    Next
                    Set nodePaths = Nothing
                    For K = 0 To .PathCount - 1
                        If .Paths(K).LayerName = .MapLayer(J).Name Then
                            If nodePaths Is Nothing Then
                                Set nodePaths = tvwProject.Nodes.Add(nodeLayer, tvwChild, , "Paths", 1)
                                nodePaths.ExpandedImage = 2
                            End If
                            Set nodePath = tvwProject.Nodes.Add(nodePaths, tvwChild, , .Paths(K).Name, 10)
                            nodePath.Tag = "PA" & CStr(I) & "," & CStr(K)
                            Set nodeSprites = Nothing
                            For L = 0 To .SpriteDefCount - 1
                                If .SpriteDefs(L).rPath Is .Paths(K) Then
                                    If nodeSprites Is Nothing Then
                                        Set nodeSprites = tvwProject.Nodes.Add(nodePath, tvwChild, , "Sprites", 1)
                                        nodeSprites.ExpandedImage = 2
                                    End If
                                    tvwProject.Nodes.Add(nodeSprites, tvwChild, , .SpriteDefs(L).Name, 5).Tag = "SD" & CStr(I) & "," & CStr(L)
                                End If
                            Next
                        End If
                    Next
                Next
            End If
            If .SpriteDefCount > 0 Then
                Set nodeSprites = tvwProject.Nodes.Add(nodeMap, tvwChild, , "Sprites", 1)
                nodeSprites.ExpandedImage = 2
                For J = 0 To .SpriteDefCount - 1
                    tvwProject.Nodes.Add(nodeSprites, tvwChild, , .SpriteDefs(J).Name, 5).Tag = "SD" & CStr(I) & "," & CStr(J)
                Next
            End If
            If .SpriteTemplateCount > 0 Then
                Set nodeTemplates = tvwProject.Nodes.Add(nodeMap, tvwChild, , "Sprite Templates", 1)
                nodeTemplates.ExpandedImage = 2
                For J = 0 To .SpriteTemplateCount - 1
                    Set nodeTemplate = tvwProject.Nodes.Add(nodeTemplates, tvwChild, , .SpriteTemplates(J).Name, 12)
                    nodeTemplate.Tag = "ST" & CStr(I) & "," & CStr(J)
                    Set nodeSprites = Nothing
                    For K = 0 To .SpriteDefCount - 1
                        If .SpriteDefs(K).Template Is .SpriteTemplates(J) Then
                            If nodeSprites Is Nothing Then
                                Set nodeSprites = tvwProject.Nodes.Add(nodeTemplate, tvwChild, , "Sprites", 1)
                                nodeSprites.ExpandedImage = 2
                            End If
                            tvwProject.Nodes.Add(nodeSprites, tvwChild, , .SpriteDefs(K).Name, 5).Tag = "SD" & CStr(I) & "," & CStr(K)
                        End If
                    Next
                Next
            End If
            If .PathCount > 0 Then
                Set nodePaths = tvwProject.Nodes.Add(nodeMap, tvwChild, , "Paths", 1)
                nodePaths.ExpandedImage = 2
                For J = 0 To .PathCount - 1
                    Set nodePath = tvwProject.Nodes.Add(nodePaths, tvwChild, , .Paths(J).Name, 10)
                    nodePath.Tag = "PA" & CStr(I) & "," & CStr(J)
                    Set nodeSprites = Nothing
                    For K = 0 To .SpriteDefCount - 1
                        If .SpriteDefs(K).rPath Is .Paths(J) Then
                            If nodeSprites Is Nothing Then
                                Set nodeSprites = tvwProject.Nodes.Add(nodePath, tvwChild, , "Sprites", 1)
                                nodeSprites.ExpandedImage = 2
                            End If
                            tvwProject.Nodes.Add(nodeSprites, tvwChild, , .SpriteDefs(K).Name, 5).Tag = "SD" & CStr(I) & "," & CStr(K)
                        End If
                    Next
                Next
            End If
        End With
    Next
    Set nodeProjElem = tvwProject.Nodes.Add(nodeProject, tvwChild, "TilesetRoot", "Tilesets", 1)
    nodeProjElem.ExpandedImage = 2
    For I = 0 To Prj.TileSetDefCount - 1
        Set nodeTileset = tvwProject.Nodes.Add(nodeProjElem, tvwChild, , Prj.TileSetDef(I).Name, 4)
        nodeTileset.Tag = "TS" & CStr(I)
        If Prj.GroupByTilesetCount(Prj.TileSetDef(I).Name) > 0 Then
            Set nodeCategories = tvwProject.Nodes.Add(nodeTileset, tvwChild, , "Categories", 1)
            nodeCategories.ExpandedImage = 2
            For J = 0 To Prj.GroupByTilesetCount(Prj.TileSetDef(I).Name) - 1
                tvwProject.Nodes.Add(nodeCategories, tvwChild, , Prj.TilesetGroupByIndex(Prj.TileSetDef(I).Name, J).Name, 8).Tag = "TC" & CStr(Prj.TilesetGroupByIndex(Prj.TileSetDef(I).Name, J).GetIndex)
            Next
        End If
        If Prj.SolidDefByTilesetCount(Prj.TileSetDef(I).Name) > 0 Then
            Set nodeSolidDefs = tvwProject.Nodes.Add(nodeTileset, tvwChild, , "Solidity Definitions", 1)
            nodeSolidDefs.ExpandedImage = 2
            For J = 0 To Prj.SolidDefByTilesetCount(Prj.TileSetDef(I).Name) - 1
                K = Prj.SolidDefIndexByTileset(Prj.TileSetDef(I).Name, J)
                tvwProject.Nodes.Add(nodeSolidDefs, tvwChild, , Prj.SolidDefsByIndex(K).Name, 11).Tag = "SO" & CStr(K)
            Next
        End If
    Next
    Set nodeProjElem = tvwProject.Nodes.Add(nodeProject, tvwChild, "MatchRoot", "Tile Matching", 1)
    nodeProject.ExpandedImage = 2
    For I = 0 To Prj.MatchDefCount - 1
        Set nodeMatch = tvwProject.Nodes.Add(nodeProjElem, tvwChild, , Prj.MatchDefs(I).Name, 6)
        nodeMatch.Tag = "MD" & CStr(I)
    Next
    Set nodeProjElem = tvwProject.Nodes.Add(nodeProject, tvwChild, "AnimRoot", "Animations", 1)
    nodeProject.ExpandedImage = 2
    For I = 0 To Prj.AnimDefCount - 1
        Set nodeAnim = tvwProject.Nodes.Add(nodeProjElem, tvwChild, , Prj.AnimDefs(I).Name, 7)
        nodeAnim.Tag = "AD" & CStr(I)
    Next
    If Prj.MediaMgr.MediaClipCount > 0 Then
        Set nodeProjElem = tvwProject.Nodes.Add(nodeProject, tvwChild, "MediaRoot", "Multimedia", 1)
    End If
    For I = 0 To Prj.MediaMgr.MediaClipCount - 1
        Set nodeMedia = tvwProject.Nodes.Add(nodeProjElem, tvwChild, , Prj.MediaMgr.Clip(I).Name, 13)
        nodeMedia.Tag = "MM" & CStr(I)
    Next
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAnimation_Click()
    frmTileAnim.Show
End Sub

Private Sub mnuCategories_Click()
    frmGroupTiles.Show
End Sub

Private Sub mnuCollisions_Click()
    frmCollisions.Show
End Sub

Private Sub mnuCtrlConfig_Click()
    frmCtrlConfig.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExport_Click()
    Dim FN As Integer
    Dim strXML As String

    On Error Resume Next

    cdlXML.DialogTitle = "Export GameDev Project as XML"
    cdlXML.Flags = cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    Err.Clear
    cdlXML.ShowSave
    If Err.Number <> 0 Then Exit Sub
    Err.Clear
    strXML = Prj2XML
    If Len(strXML) Then
        FN = FreeFile
        Open cdlXML.FileName For Output As #FN
        Print #FN, strXML
        Close #FN
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "GameDev XML Export"
    End If

End Sub

Private Sub mnuGoto_Click()
    tvwProject_DblClick
End Sub

Private Sub mnuHelpContents_Click()
    HtmlHelp Me.hwnd, App.HelpFile, HH_HELP_TOC, 0
End Sub

Private Sub mnuHelpIndex_Click()
    HtmlHelp Me.hwnd, App.HelpFile, HH_HELP_INDEX, 0
End Sub

Private Sub mnuHelpTech_Click()
    HtmlHelp Me.hwnd, App.HelpFile, HH_HELP_CONTEXT, 1001
End Sub

Private Sub mnuHelpTutorial_Click()
    HtmlHelp Me.hwnd, App.HelpFile, HH_HELP_CONTEXT, 200
End Sub

Private Sub mnuImport_Click()
    Dim FN As Integer
    Dim strXML As String

    On Error Resume Next

    cdlXML.DialogTitle = "Import GameDev Project from XML"
    cdlXML.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    Err.Clear
    cdlXML.ShowOpen
    If Err.Number <> 0 Then Exit Sub
    Err.Clear
    FN = FreeFile
    Open cdlXML.FileName For Input As #FN
    strXML = Input(LOF(FN), FN)
    Close #FN
    XML2Prj strXML
    Prj.ProjectPath = PathFromFile(cdlXML.FileName) & "\Imported.gdp"
    LoadTree
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "GameDev XML Export"
    End If
End Sub

Private Sub mnuMaps_Click()
    frmMapEdit.Show
End Sub

Private Sub mnuMatching_Click()
    frmTSMatching.Show
End Sub

Private Sub mnuMedia_Click()
    frmManageMedia.Show
End Sub

Private Sub mnuNew_Click()
    If UnloadProject Then
        Set Prj = New GameProject
    End If
    LoadTree
End Sub

Private Sub mnuOpen_Click()
    dlgProjectFile.Flags = &H100C
    On Error Resume Next
    dlgProjectFile.InitDir = GetSetting("GameDev", "Directories", "ProjectPath", App.Path)
    dlgProjectFile.ShowOpen
    If Err = 0 Then
        On Error GoTo ProjLoadErr
        If UnloadProject Then
            Prj.Load dlgProjectFile.FileName
        End If
        SaveSetting "GameDev", "Directories", "ProjectPath", Left$(dlgProjectFile.FileName, Len(dlgProjectFile.FileName) - Len(dlgProjectFile.FileTitle) - 1)
    End If
    LoadTree
    Exit Sub
    
ProjLoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuPlayer_Click()
    frmPlayer.Show
End Sub

Private Sub mnuPlayGame_Click()
    Dim ScreenDepth As Integer
    Dim xmlBackup As String
    Dim prjNameBackup As String
    Dim bDoBackup As Boolean
    Dim WarnMessage As String
    Dim TileGraphics() As StdPicture
    Dim nIdx As Integer
    Dim strErr As String
    
    On Error GoTo PlayErr
    
    bDoBackup = GetSetting("GameDev", "Options", "BackupForPlay", "1") <> 0
    
    If GetSetting("GameDev", "Options", "PlayWarn", "1") <> "0" Then
        If bDoBackup Then
            WarnMessage = "While GameDev will do its best to restore the current state of the project " & _
                          "after returning to the editor, it's generally safest not to play the game from within " & _
                          "the game editing environment in case the process must be terminated for some reason. " & _
                          "It is recommended instead that the project be saved and a shortcut be created to " & _
                          "play the game in a separate process using ""Make Shortcut"" from the File menu. " & _
                          "(This behavior can be altered from ""Options"" in the view menu.) Proceed to play?"
        Else
            WarnMessage = "The option to store the current state of the project is currently " & _
                          "turned off.  This means any changes that occur while playing the game " & _
                          "will persist in memory, and may be saved with the project permanently. " & _
                          "Proceed to play?"
        End If
        If MsgBox(WarnMessage, vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    End If

    If bDoBackup Then
        xmlBackup = Prj2XML
        prjNameBackup = Prj.ProjectPath
        If Len(xmlBackup) = 0 Then
            If MsgBox("GameDev was unable to backup the current state of the project in memory. " & _
                      "Any changes that take place to persistent objects during gameplay will " & _
                      "remain changed if the project is saved.  Proceed to play?", vbCritical + vbYesNo + vbDefaultButton2) <> vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    ScreenDepth = Val(GetSetting("GameDev", "Options", "ScreenDepth", "16"))
    If ScreenDepth <> 16 And ScreenDepth <> 24 And ScreenDepth <> 32 Then ScreenDepth = 16
    Prj.GamePlayer.Play ScreenDepth
    
    If bDoBackup Then
        ReDim TileGraphics(Prj.TileSetDefCount - 1)
        For nIdx = 0 To Prj.TileSetDefCount - 1
            If Prj.TileSetDef(nIdx).IsLoaded Then
                Set TileGraphics(nIdx) = Prj.TileSetDef(nIdx).Image
            Else
                Set TileGraphics(nIdx) = Nothing
            End If
        Next
        If Len(xmlBackup) Then
            If Not (XML2Prj(xmlBackup)) Then
                MsgBox "GameDev failed to restore the state of the project before play. " & _
                       "Be aware that saving the project will save any changes made during play.", vbExclamation
            End If
        End If
        For nIdx = 0 To Prj.TileSetDefCount - 1
            If Not (TileGraphics(nIdx) Is Nothing) Then
                Set Prj.TileSetDef(nIdx).Image = TileGraphics(nIdx)
            End If
        Next
        Prj.ProjectPath = prjNameBackup
    End If
    Exit Sub

PlayErr:
    strErr = Err.Description
    If Not (CurDisp Is Nothing) Then
        CurDisp.Close
    End If
    Set CurDisp = Nothing
    MsgBox "An error occurred while storing the project, playing the game or restoring the project: " & strErr, vbExclamation
End Sub

Private Sub mnuQuickTut_Click()
    On Error Resume Next
    Shell "Notepad.exe " & App.Path & "\Quicktut.txt", vbNormalFocus
    If Err.Number <> 0 Then MsgBox "Unable to launch Notepad.  See " & App.Path & "\QuickTut.txt for quick start tutorial.", vbExclamation
End Sub

Private Sub mnuRefresh_Click()
    LoadTree
End Sub

Private Sub mnuSave_Click()
    On Error Resume Next
    dlgProjectFile.Flags = &H880E&
    dlgProjectFile.InitDir = GetSetting("GameDev", "Directories", "ProjectPath", App.Path)
    dlgProjectFile.FileName = Prj.ProjectPath
    dlgProjectFile.ShowSave
    If Err = 0 Then
        On Error GoTo SaveErr
        Prj.Save dlgProjectFile.FileName
        SaveSetting "GameDev", "Directories", "ProjectPath", Left$(dlgProjectFile.FileName, Len(dlgProjectFile.FileName) - Len(dlgProjectFile.FileTitle) - 1)
    End If
    Exit Sub

SaveErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuShortcut_Click()
    frmShortcut.Show
End Sub

Private Sub mnuSprites_Click()
   frmSprites.Show
End Sub

Private Sub mnuTilesets_Click()
    frmTSEdit.Show
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show
End Sub

Private Sub tbrProject_ButtonClick(ByVal Button As MSComCtlLib.Button)
    Select Case Button.Key
    Case "New"
        mnuNew_Click
    Case "Open"
        mnuOpen_Click
    Case "Save"
        mnuSave_Click
    Case "Export"
        mnuExport_Click
    Case "Import"
        mnuImport_Click
    Case "Shortcut"
        mnuShortcut_Click
    Case "Play"
        mnuPlayGame_Click
    Case "Tilesets"
        mnuTilesets_Click
    Case "Maps"
        mnuMaps_Click
    Case "TileMatch"
        mnuMatching_Click
    Case "TileAnim"
        mnuAnimation_Click
    Case "Categories"
        mnuCategories_Click
    Case "Sprites"
        mnuSprites_Click
    Case "CollDef"
        mnuCollisions_Click
    Case "Player"
        mnuPlayer_Click
    Case "Media"
        mnuMedia_Click
    Case "Controls"
        mnuCtrlConfig_Click
    Case "Options"
        mnuViewOptions_Click
    Case "Help"
        mnuHelpContents_Click
    End Select
End Sub

Private Sub tvwProject_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim Idx As Integer
    Dim Idx2 As Integer
    Dim I As Integer
    
    If Len(NewString) = 0 Then
        Cancel = True
        Exit Sub
    End If
    Idx = Val(Mid$(tvwProject.SelectedItem.Tag, 3))
    I = InStr(tvwProject.SelectedItem.Tag, ",")
    If I > 0 Then
        Idx2 = Val(Mid$(tvwProject.SelectedItem.Tag, I + 1))
    End If
    
    Select Case Left$(tvwProject.SelectedItem.Tag, 2)
    Case "MD"
        For I = 0 To Prj.MatchDefCount - 1
            If Prj.MatchDefs(I).Name = NewString Then
                Cancel = True
                Exit Sub
            End If
        Next
        Prj.MatchDefs(Idx).Name = NewString
    Case "PA"
        If Prj.Maps(Idx).PathExists(NewString) Then
            Cancel = True
            Exit Sub
        End If
        Prj.Maps(Idx).Paths(Idx2).Name = NewString
    Case "AD"
        For I = 0 To Prj.AnimDefCount - 1
            If Prj.AnimDefs(I).Name = NewString Then
                Cancel = True
                Exit Sub
            End If
        Next
        Prj.AnimDefs(Idx).Name = NewString
    Case Else
        Cancel = True
    End Select
End Sub

Private Sub tvwProject_BeforeLabelEdit(Cancel As Integer)
    Select Case Left$(tvwProject.SelectedItem.Tag, 2)
    Case "MD", "PA", "AD"
    Case Else
        Cancel = True
    End Select
End Sub

Private Sub tvwProject_DblClick()
    Dim Idx As Integer
    Dim Idx2 As Integer
    Dim AD As AnimDef
    Dim MD As MatchDef
    Dim PA As Path
    Dim SD As SpriteDef
    Dim TS As TileSetDef
    Dim TC As Category
    Dim SO As SolidDef
    Dim Mp As Map
    Dim LY As Layer
    Dim ST As SpriteTemplate
    Dim MDEdit As frmMatchTile
    Dim MM As MediaClip
    Dim I As Integer
    
    On Error GoTo NavErr
    
    If tvwProject.SelectedItem Is Nothing Then Exit Sub
    Idx = Val(Mid$(tvwProject.SelectedItem.Tag, 3))
    I = InStr(tvwProject.SelectedItem.Tag, ",")
    If I > 0 Then
        Idx2 = Val(Mid$(tvwProject.SelectedItem.Tag, I + 1))
    End If
    Select Case Left$(tvwProject.SelectedItem.Tag, 2)
    Case "AD"
        Set AD = Prj.AnimDefs(Idx)
        frmTileAnim.Show
        For I = 0 To frmTileAnim.lstMaps.ListCount - 1
            If frmTileAnim.lstMaps.List(I) = AD.MapName Then
                frmTileAnim.lstMaps.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To frmTileAnim.lstLayers.ListCount - 1
            If frmTileAnim.lstLayers.List(I) = AD.LayerName Then
                frmTileAnim.lstLayers.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To frmTileAnim.lstAnimDefs.ListCount - 1
            If frmTileAnim.lstAnimDefs.List(I) = AD.Name Then
                frmTileAnim.lstAnimDefs.ListIndex = I
                Exit For
            End If
        Next
    Case "MD"
        Set MD = Prj.MatchDefs(Idx)
        
        On Error Resume Next
        If Not MD.TSDef.IsLoaded Then
            MD.TSDef.Load
        End If
        On Error GoTo NavErr
        If Not MD.TSDef.IsLoaded Then
            MsgBox "Cannot load tileset image, unable to edit tilematch."
            Exit Sub
        End If
        
        Set MDEdit = New frmMatchTile
        MDEdit.EditMatches MD
    Case "PA"
        frmSprites.Show
        
        Set PA = Prj.Maps(Idx).Paths(Idx2)
        
        For I = 0 To frmSprites.lstPaths.ListCount - 1
            If frmSprites.lstPaths.List(I) = PA.Name And _
               frmSprites.lstPaths.ItemData(I) = Idx Then
                frmSprites.lstPaths.ListIndex = I
                Exit For
            End If
        Next
    Case "SD"
        frmSprites.Show
        
        Set SD = Prj.Maps(Idx).SpriteDefs(Idx2)
        
        For I = 0 To frmSprites.lstSprites.ListCount - 1
            If frmSprites.lstSprites.List(I) = SD.Name And _
               frmSprites.lstSprites.ItemData(I) = Idx Then
                frmSprites.lstSprites.ListIndex = I
                frmSprites.cmdLoadSprite_Click
                Exit For
            End If
        Next
    Case "ST"
        frmSprites.Show
        
        Set ST = Prj.Maps(Idx).SpriteTemplates(Idx2)
        frmSprites.LoadTemplate ST
    Case "TS"
        frmTSEdit.Show
        frmTSEdit.lstTileSets.ListIndex = Idx
    Case "TC"
        frmGroupTiles.Show
        
        Set TC = Prj.GroupByIndex(Idx)
        
        For I = 0 To frmGroupTiles.cboTileset.ListCount - 1
            If frmGroupTiles.cboTileset.List(I) = TC.TSName Then
                frmGroupTiles.cboTileset.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To frmGroupTiles.cboCurGroup.ListCount - 1
            If frmGroupTiles.cboCurGroup.List(I) = TC.Name Then
                frmGroupTiles.cboCurGroup.ListIndex = I
                Exit For
            End If
        Next
    Case "SO"
        frmGroupTiles.Show
        
        Set SO = Prj.SolidDefsByIndex(Idx)
        
        For I = 0 To frmGroupTiles.cboTileset.ListCount - 1
            If frmGroupTiles.cboTileset.List(I) = SO.TSName Then
                frmGroupTiles.cboTileset.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To frmGroupTiles.cboSolidityName.ListCount - 1
            If frmGroupTiles.cboSolidityName.List(I) = SO.Name Then
                frmGroupTiles.cboSolidityName.ListIndex = I
                Exit For
            End If
        Next
    Case "MM"
        frmManageMedia.Show
        
        Set MM = Prj.MediaMgr(Idx)
        
        For I = 0 To frmManageMedia.lstMediaClips.ListCount - 1
            If frmManageMedia.lstMediaClips.List(I) = MM.Name Then
                frmManageMedia.lstMediaClips.ListIndex = I
                Exit For
            End If
        Next
    Case "MP"
        frmMapEdit.Show
        
        Set Mp = Prj.Maps(Idx)
        
        For I = 0 To frmMapEdit.lstMaps.ListCount - 1
            If frmMapEdit.lstMaps.List(I) = Mp.Name Then
                frmMapEdit.lstMaps.ListIndex = I
                Exit For
            End If
        Next
    Case "LY"
        frmMapEdit.Show
        
        Set LY = Prj.Maps(Idx).MapLayer(Idx2)
        
        For I = 0 To frmMapEdit.lstMaps.ListCount - 1
            If frmMapEdit.lstMaps.List(I) = LY.pMap.Name Then
                frmMapEdit.lstMaps.ListIndex = I
                Exit For
            End If
        Next
        
        For I = 0 To frmMapEdit.lstLayers.ListCount - 1
            If frmMapEdit.lstLayers.List(I) = LY.Name Then
                frmMapEdit.lstLayers.ListIndex = I
                Exit For
            End If
        Next
    End Select
    Exit Sub
    
NavErr:
    MsgBox Err.Description
End Sub

Private Sub tvwProject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Not (tvwProject.HitTest(X, Y) Is Nothing) Then PopupMenu mnuContext, , X, Y, mnuGoto
End Sub
