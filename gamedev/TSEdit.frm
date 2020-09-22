VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTSEdit 
   Caption         =   "Edit Tilesets"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5535
   HelpContextID   =   120
   Icon            =   "TSEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTileSetPath 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import Tile"
      Height          =   375
      Left            =   3120
      TabIndex        =   32
      ToolTipText     =   "Import a tile from another bitmap"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.OptionButton opt32Bit 
      Caption         =   "32-bit color"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton opt24Bit 
      Caption         =   "24-bit color"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton opt16Bit 
      Caption         =   "16-Bit color"
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   3720
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtTileSetName 
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      HelpContextID   =   121
      Left            =   4320
      TabIndex        =   27
      ToolTipText     =   "Edit graphics for the selected tileset"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save &All"
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      ToolTipText     =   "Save all parameters and graphics of all tilesets"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Image"
      Height          =   375
      Left            =   3120
      TabIndex        =   28
      ToolTipText     =   "Save the graphics of the selected tileset to a specified filename"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   34
      ToolTipText     =   "Close tileset editor"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4320
      TabIndex        =   31
      ToolTipText     =   "Remove the tileset from the list"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Image"
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      ToolTipText     =   "Create a new tileset based on a saved image"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      ToolTipText     =   "Commit displayed values to selected tileset"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ListBox lstTileSets 
      Height          =   1815
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtTileColumns 
      Height          =   300
      Left            =   3480
      TabIndex        =   6
      Text            =   "10"
      ToolTipText     =   "Number of tiles in each row"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtTileRows 
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Text            =   "10"
      ToolTipText     =   "Nomber of rows of tiles"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   3120
      TabIndex        =   26
      ToolTipText     =   "Create a new tileset with displayed parameters and edit its graphics"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtTileHeight 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Text            =   "32"
      ToolTipText     =   "Pixel height of each tile"
      Top             =   2640
      Width           =   495
   End
   Begin MSComCtl2.UpDown udTileWidth 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   529
      _Version        =   327681
      Value           =   32
      BuddyControl    =   "txtTileWidth"
      BuddyDispid     =   196627
      OrigLeft        =   88
      OrigTop         =   8
      OrigRight       =   101
      OrigBottom      =   33
      Max             =   64
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTileWidth 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Text            =   "32"
      ToolTipText     =   "Pixel width of each tile"
      Top             =   2280
      Width           =   495
   End
   Begin MSComCtl2.UpDown udTileHeight 
      Height          =   300
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   529
      _Version        =   327681
      Value           =   32
      BuddyControl    =   "txtTileHeight"
      BuddyDispid     =   196626
      OrigLeft        =   88
      OrigTop         =   8
      OrigRight       =   101
      OrigBottom      =   33
      Max             =   64
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udTileColumns 
      Height          =   300
      Left            =   3960
      TabIndex        =   7
      Top             =   2280
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   529
      _Version        =   327681
      Value           =   10
      BuddyControl    =   "txtTileColumns"
      BuddyDispid     =   196623
      OrigLeft        =   88
      OrigTop         =   8
      OrigRight       =   101
      OrigBottom      =   33
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udTileRows 
      Height          =   300
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   529
      _Version        =   327681
      Value           =   10
      BuddyControl    =   "txtTileRows"
      BuddyDispid     =   196624
      OrigLeft        =   88
      OrigTop         =   8
      OrigRight       =   101
      OrigBottom      =   33
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".bmp"
      Filter          =   "Windows Bitmaps (*.bmp)|*.bmp"
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen depth:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   20
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblListTilesets 
      BackStyle       =   0  'Transparent
      Caption         =   "Defined Tilesets:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTiles 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTiles"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Total tiles:"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblImageSize 
      BackStyle       =   0  'Transparent
      Caption         =   "lblImageSize"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Image size:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Rows:"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Columns:"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Height:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblTileSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Width:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmTSEdit"
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
' File: TSEdit.frm - Tileset Management Dialog
'
'======================================================================

Option Explicit

Dim WithEvents TE As TileEdit
Attribute TE.VB_VarHelpID = -1
Dim CtlPos() As POINTAPI
Dim FormWidth As Long
Dim FormHeight As Long
Dim LstWidth As Long
Dim LstHeight As Long
Dim EditTSD As TileSetDef
    
Sub Pause(L As Single)
    Dim T As Single
    T = Timer
    Do While Timer - T < L
        DoEvents
    Loop
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    On Error GoTo CreateErr
    
    If Len(txtTileSetName.Text) > 0 Then
        Set EditTSD = Prj.AddTileSet("", CInt(txtTileWidth.Text), CInt(txtTileHeight.Text), txtTileSetName.Text)
        Set TE = New TileEdit
        Set TE.Disp = New BMDXDisplay
        On Error Resume Next
        TE.Disp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
        On Error GoTo CreateErr
        TE.Disp.OpenEx , , IIf(opt24Bit.Value, 24, IIf(opt32Bit.Value, 32, 16))
        Set CurDisp = TE.Disp
        TE.Create EditTSD.TileWidth, EditTSD.TileHeight, CInt(txtTileRows.Text), CInt(txtTileColumns.Text)
        EditTSD.IsDirty = True
    Else
        MsgBox "Please enter the parameters for the tileset (including non-null name) before executing the create command"
    End If
    Exit Sub

CreateErr:
    Set TE = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdEdit_Click()
    If lstTileSets.ListIndex >= 0 Then
        Set EditTSD = Prj.TileSetDef(lstTileSets.ListIndex)
        If Not EditTSD.IsLoaded Then
            EditTSD.Load
        End If
        If EditTSD.IsLoaded Then
            Set TE = New TileEdit
            Set TE.Disp = New BMDXDisplay
            Prj.TriggerTileEdit TE
            If Not GameHost Is Nothing Then
                GameHost.RunStartScript
                If GameHost.CheckForError Then Exit Sub
            End If
            On Error Resume Next
            TE.Disp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
            TE.Disp.OpenEx , , IIf(opt24Bit.Value, 24, IIf(opt32Bit.Value, 32, 16))
            Set CurDisp = TE.Disp
            If Err.Number Then
                MsgBox Err.Description, vbOKOnly + vbExclamation
                On Error GoTo 0
                Set TE.Disp = Nothing
                Set TE = Nothing
            Else
                On Error GoTo 0
                TE.Edit EditTSD.Image, EditTSD.TileWidth, EditTSD.TileHeight
                EditTSD.IsDirty = True
            End If
        End If
    Else
        MsgBox "Please select a tileset to edit before selecting this command"
    End If
End Sub

Private Sub cmdImport_Click()
    Dim frmImp As frmTileImport
    Dim ImportedTile As StdPicture
    
    If lstTileSets.ListIndex < 0 Then
        MsgBox "Please select a tileset into which a tile should be imported before selecting this command.", vbExclamation
        Exit Sub
    End If
    
    dlgFile.Flags = &H1008
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = 0 Then
        On Error GoTo ImportErr
        Set frmImp = New frmTileImport
        With Prj.TileSetDef(lstTileSets.ListIndex)
            Set ImportedTile = frmImp.ImportTile(LoadPicture(dlgFile.FileName), .TileWidth, .TileHeight)
        End With
        Set frmImp = Nothing
        If ImportedTile Is Nothing Then Exit Sub
        Set EditTSD = Prj.TileSetDef(lstTileSets.ListIndex)
        If Not EditTSD.IsLoaded Then
            EditTSD.Load
        End If
        If EditTSD.IsLoaded Then
            Set TE = New TileEdit
            Set TE.Disp = New BMDXDisplay
            On Error Resume Next
            Me.Refresh
            DoEvents
            TE.Disp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
            TE.Disp.OpenEx , , IIf(opt24Bit.Value, 24, IIf(opt32Bit.Value, 32, 16))
            Set CurDisp = TE.Disp
            If Err.Number Then
                MsgBox Err.Description, vbOKOnly + vbExclamation
                On Error GoTo 0
                Set TE.Disp = Nothing
                Set TE = Nothing
            Else
                On Error GoTo ImportErr
                TE.Edit EditTSD.Image, EditTSD.TileWidth, EditTSD.TileHeight
                EditTSD.IsDirty = True
                DoEvents
                Set TE.TilePicture = ImportedTile
            End If
        End If
    End If
    On Error GoTo 0
    Exit Sub
    
ImportErr:
    MsgBox "Error importing tile: " & Err.Description
End Sub

Private Sub cmdLoad_Click()
    dlgFile.Flags = &H1008
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = 0 Then
        On Error GoTo 0
        Prj.AddTileSet GetRelativePath(Prj.ProjectPath, dlgFile.FileName), CInt(txtTileWidth.Text), CInt(txtTileHeight.Text), txtTileSetName.Text
        UpdateList
    End If
    On Error GoTo 0
End Sub

Private Sub cmdRemove_Click()
    If lstTileSets.ListIndex >= 0 Then
        Prj.RemoveTileSet lstTileSets.ListIndex
    End If
    UpdateList
End Sub

Private Sub cmdSave_Click()
    If lstTileSets.ListIndex >= 0 Then
        dlgFile.Flags = &H880E&
        On Error Resume Next
        dlgFile.InitDir = GetSetting("GameDev", "Directories", "TilesetPath", App.Path)
        dlgFile.ShowSave
        If Err.Number = 0 Then
            On Error GoTo 0
            With Prj.TileSetDef(lstTileSets.ListIndex)
                If Not .IsLoaded Then
                    .Load
                End If
                If .ImagePath <> GetRelativePath(Prj.ProjectPath, dlgFile.FileName) Then
                    Prj.IsDirty = True
                End If
                .ImagePath = GetRelativePath(Prj.ProjectPath, dlgFile.FileName)
                .Save
            End With
            SaveSetting "GameDev", "Directories", "TilesetPath", Left$(dlgFile.FileName, Len(dlgFile.FileName) - Len(dlgFile.FileTitle) - 1)
        End If
        On Error GoTo 0
    Else
        MsgBox "Please select a tileset to save before selecting this command"
    End If
End Sub

Private Sub cmdSaveAll_Click()
    Dim I As Integer
    
    For I = 0 To Prj.TileSetDefCount - 1
        If Prj.TileSetDef(I).ImagePath <> "" Then
            If Prj.TileSetDef(I).IsLoaded Then
                Prj.TileSetDef(I).Save
            End If
        Else
          lstTileSets.ListIndex = I
          cmdSave_Click
        End If
    Next I
End Sub

Private Sub cmdUpdate_Click()
    Dim TempDisp As BMDXDisplay
    Dim TempDC As Long
    Dim hbmpTemp As Long
    Dim hPrevObject1 As Long
    Dim hPrevObject2 As Long
    Dim rcClear As RECT
    Dim hdcOld As Long
    Dim NewPic As StdPicture
    Dim Idx As Integer, Idx2 As Integer
    
    If Not (IsNumeric(txtTileColumns.Text) And IsNumeric(txtTileRows.Text)) Then
        MsgBox "Tile Rows or Tile Columns is invalid."
        Exit Sub
    End If
    
    For Idx = 0 To Prj.MapCount - 1
        For Idx2 = 0 To Prj.Maps(Idx).LayerCount - 1
            If Prj.Maps(Idx).MapLayer(Idx2).TSDef.Name = Prj.TileSetDef(lstTileSets.ListIndex).Name Then
                If MsgBox("Tileset " & Prj.TileSetDef(lstTileSets.ListIndex).Name & " is being referenced by map """ & _
                    Prj.Maps(Idx).Name & """ layer """ & Prj.Maps(Idx).MapLayer(Idx2).Name & _
                    """.  Changing tileset parameters will most likely corrupt the map " & _
                    "and may make the map unloadable after saving.  Continue?", vbCritical + vbYesNo + vbDefaultButton2) _
                    <> vbYes Then Exit Sub
            End If
        Next
    Next
    
    If lstTileSets.ListIndex >= 0 Then
        Set EditTSD = Prj.TileSetDef(lstTileSets.ListIndex)
        With EditTSD
            On Error Resume Next
            If Not .IsLoaded Then .Load
            If Not .IsLoaded Then
                MsgBox "Unable to load image; cannot update", vbExclamation
                Exit Sub
            End If
            On Error GoTo UpdErr
        End With
                
        If CInt(Me.ScaleX(Prj.TileSetDef(lstTileSets.ListIndex).Image.Width, vbHimetric, vbPixels)) <> CInt(txtTileColumns.Text) * CInt(EditTSD.TileWidth) Or _
           CInt(Me.ScaleY(Prj.TileSetDef(lstTileSets.ListIndex).Image.Height, vbHimetric, vbPixels)) <> CInt(txtTileRows.Text) * CInt(EditTSD.TileHeight) Then
            Set TempDisp = New BMDXDisplay
            On Error Resume Next
            TempDisp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
            On Error GoTo UpdErr
            TempDisp.OpenEx , , IIf(opt24Bit.Value, 24, IIf(opt32Bit.Value, 32, 16))
            Set CurDisp = TempDisp
            TempDC = TempDisp.GetDC
            hdcOld = CreateCompatibleDC(TempDC)
            With rcClear
                .Left = 0
                .Top = 0
                .Right = CInt(txtTileColumns.Text) * EditTSD.TileWidth
                .Bottom = CInt(txtTileRows.Text) * EditTSD.TileHeight
            End With
            hbmpTemp = CreateCompatibleBitmap(TempDC, _
                rcClear.Right, rcClear.Bottom)
            TempDC = CreateCompatibleDC(TempDC)
            TempDisp.ReleaseDC
            TempDisp.Close
            Set CurDisp = Nothing
            Set TempDisp = Nothing
            hPrevObject2 = SelectObject(hdcOld, EditTSD.Image.handle)
            hPrevObject1 = SelectObject(TempDC, hbmpTemp)
            FillRect TempDC, rcClear, GetStockObject(BLACK_BRUSH)
            BitBlt TempDC, 0, 0, rcClear.Right, rcClear.Bottom, hdcOld, 0, 0, SRCCOPY
            Set NewPic = CapturePicture(TempDC, 0, 0, rcClear.Right, rcClear.Bottom)
            SelectObject TempDC, hPrevObject1
            SelectObject hdcOld, hPrevObject2
            DeleteDC hdcOld
            DeleteDC TempDC
            DeleteObject hbmpTemp
            Set EditTSD.Image = NewPic
            EditTSD.IsDirty = True
        End If
        
        With EditTSD
            .Name = txtTileSetName.Text
            .TileWidth = CInt(txtTileWidth.Text)
            .TileHeight = CInt(txtTileHeight.Text)
        End With
        Prj.IsDirty = True
    Else
        MsgBox "Please select a tileset to update before selecting this command"
    End If
    
    UpdateList
    Exit Sub

UpdErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim C As Control
    Dim WndPos As String
    
    ReDim CtlPos(Me.Controls.Count - 1)
    I = 0
    FormWidth = Me.Width
    FormHeight = Me.Height
    LstWidth = lstTileSets.Width
    LstHeight = lstTileSets.Height
    For Each C In Me.Controls
        On Error Resume Next
        CtlPos(I).X = C.Left
        CtlPos(I).Y = C.Top
        If Err.Number Then
            Err.Clear
        Else
            I = I + 1
        End If
    Next
    On Error GoTo LoadErr
    
    WndPos = GetSetting("GameDev", "Windows", "EditTileset", "")
    If WndPos <> "" Then
        Me.Move CLng(Left$(WndPos, 6)), CLng(Mid$(WndPos, 8, 6)), CLng(Mid$(WndPos, 15, 6)), CLng(Right$(WndPos, 6))
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

    ReDim Preserve CtlPos(I - 1)
    UpdateTotals
    UpdateList
    Exit Sub

LoadErr:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
    Dim C As Control
    Dim I As Integer
    
    I = 0
    On Error Resume Next
    For Each C In Me.Controls
        If TypeOf C Is CommandButton Then
            C.Left = CtlPos(I).X + Me.Width - FormWidth
        ElseIf C Is lstTileSets Then
            C.Width = LstWidth + Me.Width - FormWidth
            C.Height = LstHeight + Me.Height - FormHeight
        ElseIf C Is lblListTilesets Then
        Else
            C.Top = CtlPos(I).Y + Me.Height - FormHeight
        End If
        If Err.Number Then
            Err.Clear
        Else
            I = I + 1
        End If
    Next
    On Error GoTo 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GameDev", "Windows", "EditTileset", Format$(Me.Left, " 00000;-00000") & "," & Format$(Me.Top, " 00000;-00000") & "," & Format$(Me.Width, " 00000;-00000") & "," & Format$(Me.Height, " 00000;-00000")
End Sub

Private Sub lstTileSets_Click()
    DisplayParams Prj.TileSetDef(lstTileSets.ListIndex)
End Sub

Sub DisplayParams(TSD As TileSetDef)
    Dim P As StdPicture
    
    txtTileWidth.Text = CStr(TSD.TileWidth)
    txtTileHeight.Text = CStr(TSD.TileHeight)
    If TSD.IsLoaded Then
        txtTileColumns.Text = CStr(Me.ScaleX(TSD.Image.Width, vbHimetric, vbPixels) \ TSD.TileWidth)
        txtTileRows.Text = CStr(Me.ScaleY(TSD.Image.Height, vbHimetric, vbPixels) \ TSD.TileHeight)
    Else
        txtTileColumns.Text = ""
        txtTileRows.Text = ""
    End If
    txtTileSetName.Text = TSD.Name
    txtTileSetPath.Text = TSD.ImagePath
End Sub

Private Sub lstTileSets_DblClick()
    Dim F As New frmTSDisplay
    
    Set F.TSD = Prj.TileSetDef(lstTileSets.ListIndex)
    If F.TSD.Image Is Nothing Then
        F.TSD.Load
    End If
    F.Show
    DisplayParams Prj.TileSetDef(lstTileSets.ListIndex)
    UpdateList
End Sub

Private Sub TE_OnEditComplete()
    Set EditTSD.Image = TE.TileSetBitmap
    Set TE = Nothing
    UpdateList
    Set EditTSD = Nothing
End Sub

Private Sub txtTileColumns_Change()
    On Error Resume Next
    If udTileColumns.Value <> CInt(txtTileColumns.Text) Then
        udTileColumns.Value = CInt(txtTileColumns.Text)
    End If
    'If Err.Number Then txtTileColumns.Text = udTileColumns.Value
End Sub

Private Sub txtTileHeight_Change()
    On Error Resume Next
    If udTileHeight.Value <> CInt(txtTileHeight.Text) Then
        udTileHeight.Value = CInt(txtTileHeight.Text)
    End If
    'If Err.Number Then txtTileHeight.Text = udTileHeight.Value
End Sub

Private Sub txtTileRows_Change()
    On Error Resume Next
    If udTileRows.Value <> CInt(txtTileRows.Text) Then
        udTileRows.Value = CInt(txtTileRows.Text)
    End If
    'If Err.Number Then txtTileRows.Text = udTileRows.Value
End Sub

Private Sub txtTileWidth_Change()
    On Error Resume Next
    If udTileWidth.Value <> CInt(txtTileWidth.Text) Then
        udTileWidth.Value = CInt(txtTileWidth.Text)
    End If
    'If Err.Number Then txtTileWidth.Text = udTileWidth.Value
End Sub

Private Sub udTileColumns_Change()
    UpdateTotals
    udTileRows.Max = 256 \ udTileColumns.Value
End Sub

Private Sub udTileHeight_Change()
    UpdateTotals
End Sub

Private Sub udTileRows_Change()
    UpdateTotals
    udTileColumns.Max = 256 \ udTileRows.Value
End Sub

Private Sub udTileWidth_Change()
    udTileColumns.Max = 640 \ udTileWidth.Value
    UpdateTotals
End Sub

Sub UpdateTotals()
    On Error Resume Next
    lblImageSize.Caption = CInt(txtTileWidth.Text) * CInt(txtTileColumns.Text) & "x" & CInt(txtTileHeight.Text) * CInt(txtTileRows.Text)
    lblTiles.Caption = CInt(txtTileRows.Text * CInt(txtTileColumns.Text))
End Sub

Sub UpdateList()
    Dim I As Integer
    
    lstTileSets.Clear
    For I = 0 To Prj.TileSetDefCount - 1
        With Prj.TileSetDef(I)
            lstTileSets.AddItem IIf(.IsLoaded, "[L]", "[U]") & Prj.TileSetDef(I).Name
        End With
    Next I
End Sub
