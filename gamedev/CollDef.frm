VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCollisions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collision Definition"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   HelpContextID   =   101
   Icon            =   "CollDef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   33
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   32
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   2040
      TabIndex        =   31
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   1080
      TabIndex        =   30
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   855
   End
   Begin VB.ComboBox cboMaps 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame fraCollTest 
      Caption         =   "Collision Tests"
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
      Begin VB.CheckBox chkInvRemove 
         Caption         =   "Remove inventory after use"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboMediaClip 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3330
         Width           =   2415
      End
      Begin VB.ComboBox cboRelaventInventory 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.UpDown updOwnCount 
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   1080
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOwnCount"
         BuddyDispid     =   196619
         OrigLeft        =   2640
         OrigTop         =   1080
         OrigRight       =   2835
         OrigBottom      =   1380
         Max             =   200
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOwnCount 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   420
      End
      Begin VB.ComboBox cboOwnLack 
         Height          =   315
         ItemData        =   "CollDef.frx":0442
         Left            =   960
         List            =   "CollDef.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboCollFunc 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2955
         Width           =   2415
      End
      Begin VB.CheckBox chkBNew 
         Caption         =   "Add new B"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkANew 
         Caption         =   "Add new A"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkBTerminate 
         Caption         =   "Terminate B"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkATerminate 
         Caption         =   "Terminate A"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkBounce 
         Caption         =   "A and B swap velocities (bounce)"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CheckBox chkBStop 
         Caption         =   "B Stops"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkAStop 
         Caption         =   "A Stops"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox chkPlatform 
         Caption         =   "A rides on B"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboClassB 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cboClassA 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblMedia 
         Caption         =   "Play media clip:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblPlayerInv 
         Caption         =   "and player"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCollFunc 
         Caption         =   "Activate special function:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Terminate+new=restart; Swap+new=new instance here"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label lblClauseB 
         Caption         =   "contacts sprite B of class"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblClauseA 
         Caption         =   "When sprite A of class"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraCollisionClass 
      Caption         =   "Collision Class Names"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4695
      Begin VB.ComboBox cboClassName 
         Height          =   315
         ItemData        =   "CollDef.frx":045D
         Left            =   2160
         List            =   "CollDef.frx":045F
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblClassPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Collision Class 1 Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label lblMaps 
      BackStyle       =   0  'Transparent
      Caption         =   "Define collisions for map:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmCollisions"
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
' File: CollDef.frm - Collision Definition Dialog
'
'======================================================================

Option Explicit

Dim CurrIndex As Integer
Dim ClassIndex

Private Sub cboClassName_Change()
    Dim SSt As Integer, SLn As Integer
    On Error Resume Next
    
    SSt = cboClassName.SelStart
    SLn = cboClassName.SelLength
    Prj.Maps(cboMaps.ListIndex).CollClassName(ClassIndex) = cboClassName.Text
    cboClassA.List(ClassIndex) = cboClassName.Text
    cboClassB.List(ClassIndex) = cboClassName.Text
    cboClassName.List(ClassIndex) = cboClassName.Text
    cboClassName.ListIndex = ClassIndex
    cboClassName.SelStart = SSt
    cboClassName.SelLength = SLn
End Sub

Private Sub cboClassName_Click()
    ClassIndex = cboClassName.ListIndex
    lblClassPrompt.Caption = "Collision Class " & CStr(ClassIndex + 1) & " Name:"
End Sub

Private Sub cboMaps_Change()
    Dim I As Integer
    
    On Error GoTo LoadClassesErr
    
    cboClassName.Clear
    cboClassA.Clear
    cboClassB.Clear
    
    If cboMaps.ListIndex >= 0 Then
        With Prj.Maps(cboMaps.ListIndex)
            For I = 0 To 15
                cboClassName.AddItem .CollClassName(I)
                cboClassA.AddItem .CollClassName(I)
                cboClassB.AddItem .CollClassName(I)
            Next
    
            cboCollFunc.Clear
            cboCollFunc.AddItem "(none)"
            For I = 0 To .SpecialCount - 1
                cboCollFunc.AddItem .Specials(I).Name
            Next
            cboMediaClip.Clear
            cboMediaClip.AddItem "(none)"
            For I = 0 To Prj.MediaMgr.MediaClipCount - 1
                cboMediaClip.AddItem Prj.MediaMgr.Clip(I).Name
            Next
            cboRelaventInventory.Clear
            For I = 0 To Prj.GamePlayer.InventoryCount - 1
                cboRelaventInventory.AddItem Prj.GamePlayer.InventoryItemName(I)
            Next
    
            If .CollDefCount > 0 Then
                LoadCollDef 0
            End If
                        
        End With
        EnableClassFrame True
        cmdNew.Enabled = True
    Else
        EnableClassFrame False
        EnableTestFrame False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        cmdSave.Enabled = False
    End If
    Exit Sub
LoadClassesErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub LoadCollDef(ByVal Index As Integer)
    Dim Flags As Integer
    Dim Idx As Integer
    
    On Error GoTo LoadErr
    
    With Prj.Maps(cboMaps.ListIndex)
        If Index < .CollDefCount And Index >= 0 Then
            cboClassA.ListIndex = .CollDefs(Index).ClassA
            cboClassB.ListIndex = .CollDefs(Index).ClassB
            Flags = .CollDefs(Index).Flags
            chkPlatform.Value = IIf(Flags And eCollisionFlags.COLL_PLATFORM, vbChecked, vbUnchecked)
            chkBounce.Value = IIf(Flags And eCollisionFlags.COLL_SWAPVEL, vbChecked, vbUnchecked)
            chkAStop.Value = IIf(Flags And eCollisionFlags.COLL_ASTOP, vbChecked, vbUnchecked)
            chkBStop.Value = IIf(Flags And eCollisionFlags.COLL_BSTOP, vbChecked, vbUnchecked)
            chkATerminate.Value = IIf(Flags And eCollisionFlags.COLL_ATERM, vbChecked, vbUnchecked)
            chkBTerminate.Value = IIf(Flags And eCollisionFlags.COLL_BTERM, vbChecked, vbUnchecked)
            chkANew.Value = IIf(Flags And eCollisionFlags.COLL_ANEW, vbChecked, vbUnchecked)
            chkBNew.Value = IIf(Flags And eCollisionFlags.COLL_BNEW, vbChecked, vbUnchecked)
            cboOwnLack.ListIndex = IIf(.CollDefs(Index).InvFlags And eCollInvFlags.COLL_INV_REQUIRE, 0, 1)
            txtOwnCount.Text = .CollDefs(Index).InvUseCount
            chkInvRemove.Value = IIf(.CollDefs(Index).InvFlags And eCollInvFlags.COLL_INV_REMOVE, vbChecked, vbUnchecked)
            cboCollFunc.ListIndex = 0
            For Idx = 1 To cboCollFunc.ListCount - 1
                If cboCollFunc.List(Idx) = .CollDefs(Index).SpecialFunction Then
                    cboCollFunc.ListIndex = Idx
                    Exit For
                End If
            Next
            cboMediaClip.ListIndex = 0
            For Idx = 1 To cboMediaClip.ListCount - 1
                If cboMediaClip.List(Idx) = .CollDefs(Index).Media Then
                    cboMediaClip.ListIndex = Idx
                    Exit For
                End If
            Next
            If cboRelaventInventory.ListCount > 0 Then
                cboRelaventInventory.ListIndex = .CollDefs(Index).InvItem
            Else
                cboRelaventInventory.ListIndex = -1
            End If
            EnableTestFrame True
            If Index > 0 Then cmdPrevious.Enabled = True Else cmdPrevious.Enabled = False
            If Index < .CollDefCount - 1 Then cmdNext.Enabled = True Else cmdNext.Enabled = False
            cmdDelete.Enabled = True
            cmdSave.Enabled = True
        Else
            EnableTestFrame False
            cmdDelete.Enabled = False
            cmdPrevious.Enabled = False
            cmdNext.Enabled = False
            cmdSave.Enabled = False
        End If
    End With
    
    fraCollTest.Caption = "Collision Test " & CStr(Index + 1) & " / " & Prj.Maps(cboMaps.ListIndex).CollDefCount
    
    Exit Sub
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub SaveCollDef(Index As Integer)
    Dim Flags As Integer

    On Error GoTo SaveErr
    
    If chkPlatform.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_PLATFORM
    If chkBounce.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_SWAPVEL
    If chkAStop.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_ASTOP
    If chkATerminate.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_ATERM
    If chkANew.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_ANEW
    If chkBStop.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_BSTOP
    If chkBTerminate.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_BTERM
    If chkBNew.Value = vbChecked Then Flags = Flags Or eCollisionFlags.COLL_BNEW
    With Prj.Maps(cboMaps.ListIndex)
        .CollDefs(Index).ClassA = cboClassA.ListIndex
        .CollDefs(Index).ClassB = cboClassB.ListIndex
        .CollDefs(Index).Flags = Flags
        .CollDefs(Index).InvFlags = IIf(chkInvRemove.Value = vbChecked, eCollInvFlags.COLL_INV_REMOVE, 0) Or _
                                    IIf(cboOwnLack.ListIndex = 0, eCollInvFlags.COLL_INV_REQUIRE, 0)
        .CollDefs(Index).InvUseCount = Val(txtOwnCount.Text)
        .CollDefs(Index).InvItem = IIf(cboRelaventInventory.ListIndex >= 0, cboRelaventInventory.ListIndex, 0)
        .CollDefs(Index).SpecialFunction = IIf(cboCollFunc.ListIndex = 0, "", cboCollFunc.Text)
        .CollDefs(Index).Media = IIf(cboMediaClip.ListIndex = 0, "", cboMediaClip.Text)
        .IsDirty = True
    End With

    Exit Sub
SaveErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cboMaps_Click()
    cboMaps_Change
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DelErr
    
    Prj.Maps(cboMaps.ListIndex).RemoveCollDef CurrIndex
    If CurrIndex >= Prj.Maps(cboMaps.ListIndex).CollDefCount Then
        CurrIndex = CurrIndex - 1
    End If
    LoadCollDef CurrIndex
    
    Exit Sub
DelErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdNew_Click()
    Dim NewDef As New CollisionDef

    On Error GoTo AddErr
    
    Prj.Maps(cboMaps.ListIndex).AddCollDef NewDef
    CurrIndex = Prj.Maps(cboMaps.ListIndex).CollDefCount - 1
    LoadCollDef CurrIndex
    
    Exit Sub
AddErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    If CurrIndex < Prj.Maps(cboMaps.ListIndex).CollDefCount - 1 Then CurrIndex = CurrIndex + 1
    LoadCollDef CurrIndex
End Sub

Private Sub cmdPrevious_Click()
    If CurrIndex > 0 Then CurrIndex = CurrIndex - 1
    LoadCollDef CurrIndex
End Sub

Private Sub cmdSave_Click()
    SaveCollDef CurrIndex
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "Collisions", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    EnableClassFrame False
    EnableTestFrame False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdSave.Enabled = False
    
    For I = 0 To Prj.MapCount - 1
        cboMaps.AddItem Prj.Maps(I).Name
    Next
    Exit Sub

LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Sub EnableTestFrame(ByVal bEnable As Boolean)
    fraCollTest.Enabled = bEnable
    cboClassA.Enabled = bEnable
    cboClassB.Enabled = bEnable
    lblClauseA.Enabled = bEnable
    lblClauseB.Enabled = bEnable
    chkPlatform.Enabled = bEnable
    chkBounce.Enabled = bEnable
    chkAStop.Enabled = bEnable
    chkBStop.Enabled = bEnable
    chkATerminate.Enabled = bEnable
    chkBTerminate.Enabled = bEnable
    chkANew.Enabled = bEnable
    chkBNew.Enabled = bEnable
    lblInfo.Enabled = bEnable
    lblCollFunc.Enabled = bEnable
    cboCollFunc.Enabled = bEnable
    lblPlayerInv.Enabled = bEnable
    cboOwnLack.Enabled = bEnable
    txtOwnCount.Enabled = bEnable
    updOwnCount.Enabled = bEnable
    cboRelaventInventory.Enabled = bEnable
    chkInvRemove.Enabled = bEnable
    lblMedia.Enabled = bEnable
    cboMediaClip.Enabled = bEnable
End Sub

Sub EnableClassFrame(ByVal bEnable As Boolean)
    fraCollisionClass.Enabled = bEnable
    lblClassPrompt.Enabled = bEnable
    cboClassName.Enabled = bEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "Collisions", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub
