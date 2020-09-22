VERSION 5.00
Begin VB.Form frmTSMatching 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Tile Matching"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4410
   HelpContextID   =   122
   Icon            =   "TSMatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Assign a new name to the selected tilematch"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Delete the selected tilematch"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Edit the selected tile match"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Define new tile matching for selected tileset"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox lstTileMatches 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ListBox lstTileSets 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblMatches 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile matches defined for tileset:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblTileSets 
      BackStyle       =   0  'Transparent
      Caption         =   "Available tilesets:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmTSMatching"
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
' File: TSMatch.frm - Tile Matching Management Dialog
'
'======================================================================

Option Explicit

Private Sub cmdCreate_Click()
    Dim MD As MatchDef
    Dim strNewName As String
    Dim EditForm As frmMatchTile
    
    On Error GoTo CreateErr
    
    If lstTileSets.ListIndex < 0 Then
        MsgBox "Please select a tileset before selecting this command"
        Exit Sub
    End If
    
    strNewName = InputBox("Enter a name for the new tilematch:", "Create")
    
    If Len(strNewName) = 0 Then Exit Sub
        
    Set MD = New MatchDef
    MD.Name = strNewName
    
    Set MD.TSDef = Prj.TileSetDef(lstTileSets.List(lstTileSets.ListIndex))
    On Error Resume Next
    If Not MD.TSDef.IsLoaded Then
        MD.TSDef.Load
    End If
    On Error GoTo CreateErr
    If Not MD.TSDef.IsLoaded Then
        MsgBox "Cannot load tileset image, unable to create tilematch."
        Exit Sub
    End If
    
    Prj.AddMatch MD
    lstTileMatches.AddItem strNewName
    Set EditForm = New frmMatchTile
    EditForm.EditMatches MD
    
    Exit Sub
    
CreateErr:
    MsgBox "Error creating tilematch: " & Err.Description, vbExclamation
    
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DeleteErr
    
    If lstTileSets.ListIndex < 0 Then
        MsgBox "Please select a tileset and tilematch before selecting this command"
        Exit Sub
    End If
    
    If lstTileMatches.ListIndex < 0 Then
        MsgBox "Please select a tilematch before selecting this command"
        Exit Sub
    End If
    
    Prj.RemoveMatch lstTileMatches.List(lstTileMatches.ListIndex)
    LoadTileMatching Prj.TileSetDef(lstTileSets.List(lstTileSets.ListIndex))
    
    Exit Sub
    
DeleteErr:
    MsgBox "Error deleting tilematch: " & Err.Description
    
End Sub

Private Sub cmdEdit_Click()
    Dim MD As MatchDef
    Dim EditForm As frmMatchTile
    
    On Error GoTo EditErr
    
    If lstTileSets.ListIndex < 0 Then
        MsgBox "Please select a tileset and tilematch before selecting this command"
        Exit Sub
    End If
    
    If lstTileMatches.ListIndex < 0 Then
        MsgBox "Please select a tilematch before selecting this command"
        Exit Sub
    End If

    Set MD = Prj.MatchDefs(lstTileMatches.List(lstTileMatches.ListIndex))
    
    On Error Resume Next
    If Not MD.TSDef.IsLoaded Then
        MD.TSDef.Load
    End If
    On Error GoTo EditErr
    If Not MD.TSDef.IsLoaded Then
        MsgBox "Cannot load tileset image, unable to edit tilematch."
        Exit Sub
    End If
    
    Set EditForm = New frmMatchTile
    EditForm.EditMatches MD
    
    Exit Sub
    
EditErr:
    MsgBox "Error editing tilematch: " & Err.Description, vbExclamation
    
End Sub

Private Sub cmdRename_Click()
    Dim strNewName As String
    
    If lstTileSets.ListIndex < 0 Then
        MsgBox "Please select a tileset and a tilematch before selecting this command"
        Exit Sub
    End If
    
    With lstTileMatches
        If .ListIndex < 0 Then
            MsgBox "Please select a tilematch before selecting this command."
            Exit Sub
        End If
        strNewName = InputBox("Enter a new name for the tilematch:", "Rename", .List(.ListIndex))
        If Len(strNewName) Then
            Prj.MatchDefs(.List(.ListIndex)).Name = strNewName
            .List(.ListIndex) = strNewName
        End If
    End With
    
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim J As Integer
    Dim bRepeat As Boolean
    Dim WndPos As String
    
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "ManageMatching", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6))
        If Me.Left >= Screen.Width Then Me.Left = Screen.Width / 2
        If Me.Top >= Screen.Height Then Me.Top = Screen.Height / 2
    End If
    
    LoadTileSetList

    Do
        bRepeat = False
        For I = 0 To Prj.MatchDefCount - 1
            For J = 0 To Prj.TileSetDefCount - 1
                If Prj.MatchDefs(I).TSDef Is Prj.TileSetDef(J) Then
                    Exit For
                End If
            Next
            If J >= Prj.TileSetDefCount Then
                If MsgBox("The tileset for tilematch """ & Prj.MatchDefs(I).Name & """ is no longer available (""" & Prj.MatchDefs(I).TSDef.Name & """).  Delete """ & Prj.MatchDefs(I).Name & """?", vbExclamation + vbYesNo) = vbYes Then
                    Prj.RemoveMatch I
                    bRepeat = True
                    Exit For
                End If
            End If
        Next
    Loop While bRepeat
    Exit Sub
    
LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub LoadTileSetList()
    Dim I As Integer
    
    lstTileSets.Clear
    
    For I = 0 To Prj.TileSetDefCount - 1
        lstTileSets.AddItem Prj.TileSetDef(I).Name
    Next
    
End Sub

Private Sub LoadTileMatching(TSD As TileSetDef)
    Dim I As Integer
    
    lstTileMatches.Clear
    
    For I = 0 To Prj.MatchDefCount - 1
        With Prj.MatchDefs(I)
            If .TSDef.Name = TSD.Name Then lstTileMatches.AddItem .Name
        End With
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "ManageMatching", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000")
End Sub

Private Sub lstTileSets_Click()
    On Error Resume Next
    
    If lstTileSets.ListIndex >= 0 Then
        LoadTileMatching Prj.TileSetDef(lstTileSets.List(lstTileSets.ListIndex))
    End If
    
    If Err.Number Then MsgBox Err.Description
    
End Sub
