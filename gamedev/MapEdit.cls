VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Full screen map editor manipulates various aspects of a map.  Not directly persistable (not owned or indexed)."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MapEdit.cls - Full Screen Map Editor User Interface Class
'
'======================================================================

Option Explicit

' Increase MzWS to make a simpler maze with bigger smoother passages
' Decrease it to make a more detailed maze with twisting passages
Private MzWS As Integer
Private MzPS As Integer

Private Type Coord
    X As Integer
    Y As Integer
End Type

Private Type MenuDef
    rcBound As RECT
    Caption As String
End Type

Public theMap As Map
Attribute theMap.VB_VarUserMemId = 0
Attribute theMap.VB_VarDescription = "Reference to the Map being edited."
Public WithEvents Disp As BMDXDisplay
Attribute Disp.VB_VarHelpID = -1
Dim bHandleEvents
Dim MouseTS As IBMDXTileSet
Public MouseX As Integer
Attribute MouseX.VB_VarDescription = "Current X coordinate of the mouse."
Public MouseY As Integer
Attribute MouseY.VB_VarDescription = "Current Y coordinate of the mouse."
Public XOff As Long
Attribute XOff.VB_VarProcData = ";Behavior"
Attribute XOff.VB_VarDescription = "Current map X-Scroll offset."
Public YOff As Long
Attribute YOff.VB_VarProcData = ";Behavior"
Attribute YOff.VB_VarDescription = "Current map Y-Scroll offset."
Public XVel As Integer
Attribute XVel.VB_VarProcData = ";Behavior"
Attribute XVel.VB_VarDescription = "Current horizontal scroll velocity."
Public YVel As Integer
Attribute YVel.VB_VarProcData = ";Behavior"
Attribute YVel.VB_VarDescription = "Current vertical scroll velocity."
Dim VelTmr As Integer
Public EditLayer As Integer
Attribute EditLayer.VB_VarProcData = ";Relational"
Attribute EditLayer.VB_VarDescription = "Index of the layer being edited within theMap (the map being edited)."
Public CurTile As Byte
Attribute CurTile.VB_VarDescription = "The index of the currently selected tile."
Public bQuit As Boolean
Attribute bQuit.VB_VarDescription = "When this becomes true the editor exits."
Dim TSOffset As Long
Public EditMode As Integer ' 0 = Normal; 1 = Menu; 2 = Copy
Attribute EditMode.VB_VarProcData = ";Behavior"
Attribute EditMode.VB_VarDescription = "The current mode of the editor (Normal, Menu, Copy, TileMatch, Maze, etc)"
                           ' 3 = Copy-Selecting; 4 = Paste
                           ' 5 = Use TileMatch; 6 = Use TileMatch Eraser
                           ' 7 = Maze; 8 = Maze-Selecting
                           ' 9 = Entering Path name; 10 = Creating Path
                           ' 11 = Entering Function Name; 12 = Adding Function
Dim MenuHeight As Integer
Dim MenuItems(0 To 5) As MenuDef
Dim CopyRow As Integer
Dim CopyCol As Integer
Dim CopyData() As Byte
Public MenuMode As Integer ' 0 = Tiles+Menu
Attribute MenuMode.VB_VarDescription = "Menu showing: 0 = Tiles + Top Menu, 1 = TileMatch, 2 = Path, 3 = Maze."
                           ' 1 = Showing TileMatch List
                           ' 2 = Showing Path Menu
                           ' 3 = Sowing Maze Menu
Dim PathMenu(0 To 3) As MenuDef
Dim MazeMenu(0 To 1) As MenuDef
Dim UseMatchDef As MatchDef
Public GeneratingMaze As Boolean
Attribute GeneratingMaze.VB_VarProcData = ";Behavior"
Attribute GeneratingMaze.VB_VarDescription = "True if a maze is currently being generated."
Dim ReadString As String
Dim ReadPos As Integer
Public NewPath As Path
Attribute NewPath.VB_VarDescription = "The Path currently being created."
Dim FlashTimer As Integer
Public bShowAllPaths As Boolean
Attribute bShowAllPaths.VB_VarDescription = "Set to True to display all paths on the layer being edited."
Public DisplayMessage As String
Attribute DisplayMessage.VB_VarDescription = "If this contains a string, it is being displayed until enter is pressed."
Public bPathGrid As Boolean
Attribute bPathGrid.VB_VarProcData = ";Behavior"
Attribute bPathGrid.VB_VarDescription = "True if the grid is turned on for creating path points."
Public NewSpec As SpecialFunction
Attribute NewSpec.VB_VarDescription = "The SpecialFunction currently being created."

Event OnEditInit()
Attribute OnEditInit.VB_Description = "Fired after all initialization, right before entering the main editor loop."
Event OnEditComplete()
Attribute OnEditComplete.VB_Description = "Fired after all the clean-up is done after exiting the editor."
Event OnBeforeMoveSprites()
Attribute OnBeforeMoveSprites.VB_Description = "Fired right before sprite movements are processed."
Event OnAfterMoveSprites()
Attribute OnAfterMoveSprites.VB_Description = "Fired right after sprite movements are processed."
Event OnAfterDraw()
Attribute OnAfterDraw.VB_Description = "Fired after the screen is drawn (good for an event that processes every loop."
Event OnKeyPress(ByVal KeyAscii As Integer)
Attribute OnKeyPress.VB_Description = "Fired whenever a keypress occurs on the editor."
Event OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseMove.VB_Description = "Fired when mouse moves accross full screen editor display."
Event OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseDown.VB_Description = "Fired when mouse os pressed on the editor."
Event OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseUp.VB_Description = "Fired when mouse is released in editor."

Private Sub InitMenu()
    With MenuItems(0)
        With .rcBound
            .Left = 0
            .Top = 0
            .Right = 45
            .Bottom = 15
        End With
        .Caption = "Close"
    End With
    With MenuItems(1)
        With .rcBound
            .Left = 50
            .Top = 0
            .Right = 95
            .Bottom = 15
        End With
        .Caption = "Copy"
    End With
    With MenuItems(2)
        With .rcBound
            .Left = 100
            .Top = 0
            .Right = 175
            .Bottom = 15
        End With
        .Caption = "Tilematch"
    End With
    With MenuItems(3)
        With .rcBound
            .Left = 180
            .Top = 0
            .Right = 240
            .Bottom = 15
        End With
        .Caption = "Maze"
    End With
    With MenuItems(4)
        With .rcBound
            .Left = 245
            .Top = 0
            .Right = 315
            .Bottom = 15
        End With
        .Caption = "Paths"
    End With
    With MenuItems(5)
        With .rcBound
            .Left = 320
            .Top = 0
            .Right = 400
            .Bottom = 15
        End With
        .Caption = "Add Special"
    End With
    MenuHeight = 16
    With PathMenu(0)
        With .rcBound
            .Left = MenuItems(4).rcBound.Left
            .Right = MenuItems(4).rcBound.Right
            .Top = MenuItems(4).rcBound.Bottom
            .Bottom = .Top + 9
        End With
        .Caption = "New Path"
    End With
    With PathMenu(1)
        With .rcBound
            .Left = MenuItems(4).rcBound.Left
            .Right = MenuItems(4).rcBound.Right
            .Top = MenuItems(4).rcBound.Bottom + 10
            .Bottom = .Top + 9
        End With
        .Caption = "Show All"
    End With
    With PathMenu(2)
        With .rcBound
            .Left = MenuItems(4).rcBound.Left
            .Right = MenuItems(4).rcBound.Right
            .Top = MenuItems(4).rcBound.Bottom + 20
            .Bottom = .Top + 9
        End With
        .Caption = "Reset Sprites"
    End With
    With PathMenu(3)
        With .rcBound
            .Left = MenuItems(4).rcBound.Left
            .Right = MenuItems(4).rcBound.Right
            .Top = MenuItems(4).rcBound.Bottom + 30
            .Bottom = .Top + 9
        End With
        .Caption = "Toggle Grid"
    End With
    With MazeMenu(0)
        With .rcBound
            .Left = MenuItems(3).rcBound.Left
            .Right = MenuItems(3).rcBound.Right
            .Top = MenuItems(3).rcBound.Bottom
            .Bottom = .Top + 9
        End With
        .Caption = "Normal"
    End With
    With MazeMenu(1)
        With .rcBound
            .Left = MenuItems(3).rcBound.Left
            .Right = MenuItems(3).rcBound.Right
            .Top = MenuItems(3).rcBound.Bottom + 10
            .Bottom = .Top + 9
        End With
        .Caption = "Large Scale"
    End With
    
End Sub

Private Sub Class_Initialize()
    CurTile = 1
    InitMenu
    MzWS = 6
    MzPS = 1
End Sub

Private Sub Disp_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    On Error GoTo KeyErr
    
    If EditMode = 9 Or EditMode = 11 Then
        Select Case KeyCode
        Case vbKeyBack
            If ReadPos > 0 Then
                ReadString = Left$(ReadString, ReadPos - 1) & Mid$(ReadString, ReadPos + 1)
                ReadPos = ReadPos - 1
            End If
        Case vbKeyDelete
            If ReadPos < Len(ReadString) Then
                ReadString = Left$(ReadString, ReadPos) & Mid$(ReadString, ReadPos + 2)
            End If
        Case vbKeyLeft
            If ReadPos > 0 Then ReadPos = ReadPos - 1
        Case vbKeyRight
            If ReadPos < Len(ReadString) Then ReadPos = ReadPos + 1
        End Select
    Else
        If GetKeyState(vbKeyScrollLock) = 0 Then
            With theMap.MapLayer(EditLayer).Tileset
                If KeyCode = vbKeyUp Then YOff = YOff - Int(theMap.ViewHeight / .TileHeight) * .TileHeight
                If KeyCode = vbKeyDown Then YOff = YOff + Int(theMap.ViewHeight / .TileHeight) * .TileHeight
                If KeyCode = vbKeyLeft Then XOff = XOff - Int(theMap.ViewWidth / .TileWidth) * .TileWidth
                If KeyCode = vbKeyRight Then XOff = XOff + Int(theMap.ViewWidth / .TileWidth) * .TileWidth
                If KeyCode = vbKeyEnd Then XOff = theMap.MapWidth - theMap.ViewWidth: YOff = theMap.MapHeight - theMap.ViewHeight
                If KeyCode = vbKeyHome Then XOff = 0: YOff = 0
                If XOff < 0 Then XOff = 0
                If XOff > theMap.MapWidth - theMap.ViewWidth Then
                    XOff = theMap.MapWidth - theMap.ViewWidth
                End If
                If YOff < 0 Then YOff = 0
                If YOff > theMap.MapHeight - theMap.ViewHeight Then
                    YOff = theMap.MapHeight - theMap.ViewHeight
                End If
            End With
        End If
        If KeyCode = vbKeyS And (Shift And vbCtrlMask) > 0 Then
            DisplayMessage = SaveScreenShot
        End If
    End If
    Exit Sub
    
KeyErr:
    Disp.Close
    MsgBox Err.Description
    bQuit = True
End Sub

Private Sub Disp_KeyPress(ByVal KeyAscii As Integer)
    On Error GoTo KeyErr
    
    If EditMode = 9 Then
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = Asc(" ") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
            ReadString = Left$(ReadString, ReadPos) & Chr$(KeyAscii) & Mid$(ReadString, ReadPos + 1)
            ReadPos = ReadPos + 1
        End If
        If KeyAscii = 27 Then EditMode = 0
        If KeyAscii = 13 Then
            If theMap.PathExists(ReadString) Then
                DisplayMessage = "Path """ & ReadString & """ already exists"
                EditMode = 0
                Set NewPath = Nothing
            Else
                EditMode = 10
                Set NewPath = New Path
                NewPath.Name = ReadString
                NewPath.LayerName = theMap.MapLayer(EditLayer).Name
            End If
        End If
    ElseIf EditMode = 10 Or EditMode = 12 Then
        If KeyAscii = 27 Then
            EditMode = 0
            Set NewPath = Nothing
            Set NewSpec = Nothing
        End If
    ElseIf EditMode = 11 Then
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = Asc(" ") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
            ReadString = Left$(ReadString, ReadPos) & Chr$(KeyAscii) & Mid$(ReadString, ReadPos + 1)
            ReadPos = ReadPos + 1
        End If
        If KeyAscii = 27 Then EditMode = 0
        If KeyAscii = 13 Then
            If theMap.SpecialExists(ReadString) Then
                DisplayMessage = "Special function """ & ReadString & """ already exists"
                EditMode = 0
                Set NewSpec = Nothing
            Else
                EditMode = 12
                Set NewSpec = New SpecialFunction
                NewSpec.Name = ReadString
                NewSpec.TileRight = -1
                NewSpec.LayerIndex = EditLayer
            End If
        End If
    Else
        If KeyAscii = 27 Then bQuit = True
        If KeyAscii = 13 Then DisplayMessage = ""
    End If
    
    RaiseEvent OnKeyPress(KeyAscii)
    
    Exit Sub
    
KeyErr:
    Disp.Close
    MsgBox Err.Description
    bQuit = True
End Sub

Private Sub Disp_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If Not bHandleEvents Then Exit Sub
    Dim Col As Integer, Row As Integer
    Dim I As Integer, J As Integer
    Dim MatchBlk As Variant
    Dim ErrStr As String
    
    On Error GoTo ChangeErr
    
    If Button = 1 Then
        MouseX = X
        MouseY = Y
        If EditMode = 0 Then
            theMap.ChangeScreenTile XOff, YOff, MouseX, MouseY, EditLayer, CurTile
            DrawScreen
        ElseIf EditMode = 2 Or EditMode = 7 Then
            If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, Col, Row) Then
                CopyCol = Col
                CopyRow = Row
                EditMode = EditMode + 1
            End If
        ElseIf EditMode = 4 Then
            If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, Col, Row) Then
                For J = Row To Row + UBound(CopyData, 2)
                    If J >= theMap.MapLayer(EditLayer).Rows Then Exit For
                    For I = Col To Col + UBound(CopyData, 1)
                        If I >= theMap.MapLayer(EditLayer).Columns Then Exit For
                        theMap.MapLayer(EditLayer).Data.TileValue(I, J) = CopyData(I - Col, J - Row)
                    Next
                Next
                theMap.IsDirty = True
            End If
            DrawScreen
        ElseIf EditMode = 5 Or EditMode = 6 Then
            With theMap.MapLayer(EditLayer)
                If theMap.ScreenToTile(XOff, YOff, MouseX - .Tileset.TileWidth / 2, MouseY - .Tileset.TileHeight / 2, EditLayer, Col, Row) Then
                    MatchBlk = Match2x2(Col, Row, EditMode = 6)
                    For J = Row - 1 To Row + 2
                        For I = Col - 1 To Col + 2
                            If I >= 0 And I < theMap.MapLayer(EditLayer).Columns And J >= 0 And J < theMap.MapLayer(EditLayer).Rows Then
                                .Data.TileValue(I, J) = MatchBlk(I - Col + 1, J - Row + 1)
                            End If
                        Next
                    Next
                    theMap.IsDirty = True
                End If
            End With
        ElseIf EditMode = 10 Then
            With theMap.MapLayer(EditLayer)
                If bPathGrid Then
                    If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, Col, Row) Then
                        NewPath.AddPoint CLng(Col) * .TSDef.TileWidth, CLng(Row) * .TSDef.TileHeight
                    End If
                Else
                    NewPath.AddPoint XOff * .XScrollRate + MouseX - theMap.ViewLeft, _
                                     YOff * .YScrollRate + MouseY - theMap.ViewTop
                End If
            End With
        ElseIf EditMode = 12 Then
            theMap.ScreenToTile XOff, YOff, MouseX, MouseY, EditLayer, Col, Row
            NewSpec.TileLeft = Col
            NewSpec.TileTop = Row
            NewSpec.TileRight = Col
            NewSpec.TileBottom = Row
        End If
    End If
    If Button = 2 Then
        If EditMode = 10 Then
            theMap.AddPath NewPath
            Set NewPath = Nothing
            EditMode = 0
        Else
            EditMode = 1
        End If
    End If
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
    Exit Sub
    
ChangeErr:
    ErrStr = Err.Description
    Disp.Close
    MsgBox ErrStr
    bQuit = True
End Sub

Private Sub Disp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Col As Integer, Row As Integer
    
    If Not bHandleEvents Then Exit Sub
    On Error GoTo MoveErr
    
    MouseX = X
    MouseY = Y
    
    If Button = 1 And EditMode = 0 Then
        theMap.ChangeScreenTile XOff, YOff, MouseX, MouseY, EditLayer, CurTile
        DrawScreen
    End If
    If (EditMode = 5 Or EditMode = 6) And Button = 1 Then
        Disp_MouseDown 1, Shift, X, Y
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
    
    Exit Sub
    
MoveErr:
    Disp.Close
    MsgBox Err.Description
    bQuit = True
    
End Sub

Sub Edit(EditMap As Map, ByVal Layer As Integer)
Attribute Edit.VB_Description = "Run the editor loop that executes until editing is complete (bQuit)."
    Dim hDC As Long
    Dim rcSel As RECT
    Dim CopyToCol As Integer
    Dim CopyToRow As Integer
    Dim ErrStr As String

    On Error GoTo EditErr
            
    If Disp Is Nothing Then Err.Raise vbObjectError, , "No display"
    Set theMap = EditMap
    Set theMap.Disp = Disp
    theMap.LoadTiles
    theMap.LoadTileAnims
    theMap.InitSprites
    If Prj.MediaMgr.ClipExists(theMap.BackgroundMusic) Then
        Prj.MediaMgr.Clip(theMap.BackgroundMusic).Play
    End If

    Set MouseTS = Disp.CreateTileSet(GetMouseBmp(), 17, 17)
    EditLayer = Layer
    Prj.GamePlayer.ReIndexTilesetRefs
    Set Prj.GamePlayer.rMap = theMap
    Set Prj.GamePlayer.Disp = Disp
    Prj.GamePlayer.InitPlayerSprite
    
    bQuit = False
    theMap.bDisablePlayerEdit = (Val(GetSetting("GameDev", "Options", "DisablePlayerEdit", "0")) <> 0)
    RaiseEvent OnEditInit
    bHandleEvents = True
    Do
        If theMap.bActive Then Disp.Cls
        DoEvents
        If EditMode <> 1 Then
            If GeneratingMaze Then GenerateMaze
            If GetKeyState(vbKeyScrollLock) And Not (Prj.GamePlayer.PlayerSprite Is Nothing) Then
                If Not GameHost Is Nothing Then
                    GameHost.RunStartScript
                    If GameHost.CheckForError Then Exit Do
                End If
                XOff = Prj.GamePlayer.MapScrollX
                YOff = Prj.GamePlayer.MapScrollY
            Else
                CheckScroll
            End If
            DrawScreen
            If theMap.bActive Then
                RaiseEvent OnAfterDraw
                If Prj.GamePlayer.bEnableJoystick Then
                    Prj.GamePlayer.PollJoystick JOYSTICKID1
                    If Prj.GamePlayer.JSValid Then Prj.GamePlayer.JoyMove
                End If
                Prj.GamePlayer.CtlActions = Prj.GamePlayer.KeyActions Or Prj.GamePlayer.JoyActions
                RaiseEvent OnBeforeMoveSprites
                theMap.AdvanceSprites
                RaiseEvent OnAfterMoveSprites
                If GetKeyState(vbKeyScrollLock) And Not (Prj.GamePlayer.PlayerSprite Is Nothing) Then
                    Prj.GamePlayer.Touch.Move Prj.GamePlayer.PlayerSprite.X, Prj.GamePlayer.PlayerSprite.Y
                    Prj.GamePlayer.Touch.TouchTest Prj.GamePlayer.PlayerSprite, Prj.GamePlayer.InterestingTiles, Prj.GamePlayer.FirstTouchTiles
                    If Not (theMap Is Prj.GamePlayer.rMap) Then
                        Set theMap = Prj.GamePlayer.rMap
                        EditLayer = 0
                        CurTile = 0
                    End If
                    Prj.GamePlayer.AutoScrollMap
                    Prj.GamePlayer.DrawInventory
                End If
                theMap.CheckSpriteCollisions
                If EditMode = 3 Or EditMode = 8 Then
                    If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, CopyToCol, CopyToRow) Then
                        hDC = Disp.GetDC
                        With theMap.MapLayer(EditLayer)
                            rcSel.Left = theMap.ViewLeft + CLng(CopyCol) * .TSDef.TileWidth - XOff * theMap.MapLayer(EditLayer).XScrollRate
                            rcSel.Top = theMap.ViewTop + CLng(CopyRow) * .TSDef.TileHeight - YOff * theMap.MapLayer(EditLayer).YScrollRate
                            rcSel.Right = theMap.ViewLeft + CLng(CopyToCol) * .TSDef.TileWidth - XOff * theMap.MapLayer(EditLayer).XScrollRate + .TSDef.TileWidth
                            rcSel.Bottom = theMap.ViewTop + CLng(CopyToRow) * .TSDef.TileHeight - YOff * theMap.MapLayer(EditLayer).YScrollRate + .TSDef.TileHeight
                        End With
                        DrawFocusRect hDC, rcSel
                        Disp.ReleaseDC
                    End If
                End If
                Disp.DrawTile MouseTS, 2, MouseX, MouseY, 0, 0, 640, 480, True
            End If
        ElseIf EditMode = 1 Then
            PaintMenu
            If theMap.bActive Then
                RaiseEvent OnAfterDraw
                If MouseY > 464 Then
                    TSOffset = TSOffset + 8
                    If TSOffset > Disp.HimetricToPixelY(theMap.MapLayer(EditLayer).TSDef.Image.Height) - 480 + MenuHeight Then
                        TSOffset = Disp.HimetricToPixelY(theMap.MapLayer(EditLayer).TSDef.Image.Height) - 480 + MenuHeight
                    End If
                End If
                If MouseY < 16 Then
                    TSOffset = TSOffset - 8
                End If
                If TSOffset < 0 Then TSOffset = 0
                Disp.DrawTile MouseTS, 2, MouseX, MouseY, 0, 0, 640, 480, True
            End If
        End If
        Prj.MediaMgr.RenderActiveMedia
        Disp.Flip
    Loop Until bQuit
    bHandleEvents = False
    
    Set Prj.GamePlayer.Disp = Nothing
    RemoveMouseTS
    Set MouseTS = Nothing
    theMap.Disp.Close
    If Not GameHost Is Nothing Then GameHost.CheckForError
    Set theMap = Nothing
    Prj.MediaMgr.ModalFadeOutAll
    Prj.MediaMgr.UnloadAll
    RaiseEvent OnEditComplete
    If Not GameHost Is Nothing Then
        GameHost.RunStartScript
        GameHost.CheckForError
    End If
    Exit Sub

EditErr:
    ErrStr = Err.Description
    bQuit = True
    Disp.Close
    Set Disp = Nothing
    Set CurDisp = Nothing
    MsgBox "Error editing map: " & ErrStr, vbCritical
    If Not GameHost Is Nothing Then GameHost.CheckForError
End Sub

Sub RemoveMouseTS()
Attribute RemoveMouseTS.VB_Description = "Clean up/release the mouse pointer graphics."
    Dim I As Integer

    For I = 0 To Disp.TileSets.Count - 1
        If MouseTS Is Disp.TileSets(I) Then
            Disp.TileSets.Remove I
        End If
    Next

End Sub

Private Sub Disp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim CopyToCol As Integer
    Dim CopyToRow As Integer
    Dim I As Integer, J As Integer
    
    If Button = 2 Then
        If EditMode = 1 Then
            ActivateMenu X, Y
        End If
    End If
    If Button = 1 Then
        If EditMode = 3 Then
            If theMap.ScreenToTile(XOff, YOff, X, Y, EditLayer, CopyToCol, CopyToRow) Then
                If CopyToCol >= CopyCol And CopyToRow >= CopyRow Then
                    ReDim CopyData(CopyToCol - CopyCol, CopyToRow - CopyRow) As Byte
                    For J = CopyRow To CopyToRow
                        For I = CopyCol To CopyToCol
                            CopyData(I - CopyCol, J - CopyRow) = theMap.MapLayer(EditLayer).Data.TileValue(I, J)
                        Next
                    Next
                End If
            End If
            EditMode = 4
        End If
        If EditMode = 8 Then
            If theMap.ScreenToTile(XOff, YOff, X, Y, EditLayer, CopyToCol, CopyToRow) Then
                If CopyToCol >= CopyCol And CopyToRow >= CopyRow Then
                    GenerateMaze True, CopyCol, CopyRow, CopyToCol, CopyToRow
                End If
                theMap.IsDirty = True
            End If
            EditMode = 0
        End If
        If EditMode = 12 Then
            theMap.AddSpecial NewSpec
            Set NewSpec = Nothing
            EditMode = 0
        End If
    End If
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
End Sub

Private Sub GenerateMaze(Optional ByVal Init As Boolean = False, Optional ByVal InitLeft As Integer, Optional ByVal InitTop As Integer, Optional ByVal InitRight As Integer, Optional ByVal InitBottom As Integer)
    Static X As Integer, Y As Integer
    Static Left As Long, Top As Long
    Static Right As Integer, Bottom As Integer
    Static Stk() As Coord
    Static StkTop As Integer
    Static StkSiz As Integer
    Dim X2 As Integer, Y2 As Integer
    Dim X3 As Integer, Y3 As Integer
    Dim Cnt As Integer
    Dim d As Integer
    Dim Tst As Byte
    Static L As Integer
    Dim T As Integer
    Dim X4 As Integer, Y4 As Integer
    Dim SolidVal As Byte
    Dim ArMatch As Variant
    
    If UseMatchDef Is Nothing Then
        SolidVal = CurTile
    Else
        On Error Resume Next
        SolidVal = UseMatchDef.TileMatches.MatchGroup(6).GetArray(0)
        If Err.Number Then SolidVal = CurTile
        On Error GoTo 0
    End If
    
    If Init Then
        Randomize Timer
        StkTop = 0
        StkSiz = 0
    
        If InitBottom < InitTop Then
            Top = InitBottom
            Bottom = InitTop
        Else
            Top = InitTop
            Bottom = InitBottom
        End If
        
        If InitRight < InitLeft Then
            Left = InitRight
            Right = InitLeft
        Else
            Left = InitLeft
            Right = InitRight
        End If
        
        With theMap.MapLayer(EditLayer).Data
            For Y = InitTop To InitBottom
                For X = InitLeft To InitRight
                    If X = InitLeft Then
                        If Not (UseMatchDef Is Nothing) Then
                            ArMatch = UseMatchDef.TileMatches.MatchGroup(5).GetArray
                            .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                        Else
                            .TileValue(X, Y) = SolidVal
                        End If
                    ElseIf X = InitRight Then
                        If Not (UseMatchDef Is Nothing) Then
                            ArMatch = UseMatchDef.TileMatches.MatchGroup(7).GetArray
                            .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                        Else
                            .TileValue(X, Y) = SolidVal
                        End If
                    Else
                        .TileValue(X, Y) = SolidVal
                    End If
                    If Y = InitTop Then
                        If X = InitLeft Then
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(0).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        ElseIf X = InitRight Then
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(2).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        Else
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(1).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        End If
                    ElseIf Y = InitBottom Then
                        If X = InitLeft Then
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(10).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        ElseIf X = InitRight Then
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(12).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        Else
                            If Not (UseMatchDef Is Nothing) Then
                                ArMatch = UseMatchDef.TileMatches.MatchGroup(11).GetArray
                                .TileValue(X, Y) = ArMatch(Int(Rnd(1) * (UBound(ArMatch) - LBound(ArMatch) + 1)) + LBound(ArMatch))
                            End If
                        End If
                    End If
                Next
            Next
        End With

        X = InitLeft + MzPS + 2
        Y = InitTop + MzPS + 2
        
        ReDim Stk((Right - Left) * (Bottom - Top) / MzWS) As Coord
        
        L = 0
        GeneratingMaze = True
        
        Exit Sub
    End If
        
    d = Int(Rnd(1) * 4) * 2
    Cnt = 0
    Do
        Do
            Select Case d
            Case 0: X2 = X - MzWS: Y2 = Y
            Case 1: X2 = X - MzWS: Y2 = Y - MzWS
            Case 2: X2 = X: Y2 = Y - MzWS
            Case 3: X2 = X + MzWS: Y2 = Y - MzWS
            Case 4: X2 = X + MzWS: Y2 = Y
            Case 5: X2 = X + MzWS: Y2 = Y + MzWS
            Case 6: X2 = X: Y2 = Y + MzWS
            Case 7: X2 = X - MzWS: Y2 = Y + MzWS
            End Select
            d = (d + 1) Mod 8
            Cnt = Cnt + 1
        Loop Until (X2 > Left + MzPS + 1 And X2 < Right - MzPS - 1 And Y2 > Top + MzPS + 1 And Y2 < Bottom - MzPS - 1) Or Cnt >= 8
        
        If Cnt < 8 Then
            Tst = 0
            With theMap.MapLayer(EditLayer).Data
                For T = MzWS / 2 To MzWS
                    If .TileValue((CLng(X2) * T + CLng(X) * (MzWS - T)) / MzWS, (CLng(Y2) * T + CLng(Y) * (MzWS - T)) / MzWS) = 0 Then
                        Tst = 1
                        Exit For
                    End If
                Next T
            End With
        Else
            Tst = 1
        End If
    Loop Until Tst = 0 Or Cnt >= 8
    If Tst = 0 And L < 10 Then
        L = L + 1
        Push X, Y, Stk, StkTop, StkSiz
        With theMap.MapLayer(EditLayer).Data
            For d = 0 To MzWS
                X3 = (CLng(X2) * d + CLng(X) * (MzWS - d)) / MzWS
                Y3 = (CLng(Y2) * d + CLng(Y) * (MzWS - d)) / MzWS
                For Y4 = Y3 - MzPS To Y3 + MzPS
                    For X4 = X3 - MzPS To X3 + MzPS
                        If (Abs(X4 - X3) + Abs(Y4 - Y3) <= MzPS) And X4 > Left And X4 < Right And Y4 > Top And Y4 < Bottom Then
                            .TileValue(X4, Y4) = 0
                        End If
                    Next
                Next
            Next
            If Not (UseMatchDef Is Nothing) Then
                For Y4 = Min(Y, Y2) - MzPS - 1 To Max(Y, Y2) + MzPS + 1
                    For X4 = Min(X, X2) - MzPS - 1 To Max(X, X2) + MzPS + 1
                        If X4 > Left And X4 < Right And Y4 > Top And Y4 < Bottom Then
                            If .TileValue(X4, Y4) > 0 Then
                                If (.TileValue(X4 - 1, Y4) > 0) + _
                                   (.TileValue(X4, Y4 - 1) > 0) + _
                                   (.TileValue(X4 + 1, Y4) > 0) + _
                                   (.TileValue(X4, Y4 + 1) > 0) >= -1 Then
                                    .TileValue(X4, Y4) = 0
                                End If
                            End If
                        End If
                    Next
                Next
                For Y4 = Min(Y, Y2) - MzPS - 1 To Max(Y, Y2) + MzPS + 1
                    For X4 = Min(X, X2) - MzPS - 1 To Max(X, X2) + MzPS + 1
                        If X4 > Left And X4 < Right And Y4 > Top And Y4 < Bottom Then
                            If .TileValue(X4, Y4) > 0 Then
                                .TileValue(X4, Y4) = UseMatchDef.MatchMiddleTile( _
                                                    .TileValue(X4 - 1, Y4 - 1), _
                                                    .TileValue(X4, Y4 - 1), _
                                                    .TileValue(X4 + 1, Y4 - 1), _
                                                    .TileValue(X4 - 1, Y4), _
                                                    .TileValue(X4 + 1, Y4), _
                                                    .TileValue(X4 - 1, Y4 + 1), _
                                                    .TileValue(X4, Y4 + 1), _
                                                    .TileValue(X4 + 1, Y4 + 1))
                            End If
                        End If
                    Next
                Next
            End If
        End With
        X = X2: Y = Y2
    Else
        L = 0
        If StkSiz > 0 Then
            Pop X, Y, Stk, StkTop, StkSiz
        Else
            StkSiz = -1
        End If
    End If
    
    If StkSiz < 0 Then
        GeneratingMaze = False
    End If
    
End Sub

Private Function Min(ByVal A, ByVal B)
    If A < B Then Min = A Else Min = B
End Function

Private Function Max(ByVal A, ByVal B)
    If A > B Then Max = A Else Max = B
End Function

Private Sub Push(ByVal X As Integer, ByVal Y As Integer, ByRef Stk() As Coord, ByRef StkTop As Integer, ByRef StkSiz As Integer)
    Stk(StkTop).X = X
    Stk(StkTop).Y = Y
    StkTop = ((StkTop + 1) Mod (UBound(Stk) - LBound(Stk) + 1)) + LBound(Stk)
    StkSiz = StkSiz + 1
    
    
    If StkSiz > UBound(Stk) Then Stop
End Sub

Private Sub Pop(ByRef X As Integer, ByRef Y As Integer, ByRef Stk() As Coord, ByRef StkTop As Integer, ByRef StkSiz As Integer)
    Dim Elem As Integer
    
    If StkSiz < 1 Then Stop
    Elem = StkTop - StkSiz
    If Elem < LBound(Stk) Then Elem = Elem + (UBound(Stk) - LBound(Stk) + 1)
    
    X = Stk(Elem).X
    Y = Stk(Elem).Y
    StkSiz = StkSiz - 1
    
End Sub

Private Sub Disp_TileSetLost(ByVal Index As Integer)
    If Not bHandleEvents Then Exit Sub
    On Error Resume Next
    If Disp.TileSets(Index) Is MouseTS Then
        MouseTS.Restore
        MouseTS.PaintPicture GetMouseBmp(), 0, 0
    End If
End Sub

Sub CheckScroll()
Attribute CheckScroll.VB_Description = "Scroll the map according to the mouse position and scrolling inertia."
    Dim bCenterX As Boolean
    Dim bCenterY As Boolean
    
    bCenterX = True
    bCenterY = True
    
    VelTmr = (VelTmr + 1) Mod 2
    If (VelTmr = 0) Or (XVel = 0 And YVel = 0) Then
        If MouseX < 16 Then
            If XVel > (MouseX - 16) Then
                XVel = XVel - 1
            ElseIf XVel < (MouseX - 16) Then
                XVel = XVel + 1
            End If
            bCenterX = False
        End If
        If MouseX > 623 Then
            If XVel > (MouseX - 623) Then
                XVel = XVel - 1
            ElseIf XVel < (MouseX - 623) Then
                XVel = XVel + 1
            End If
            bCenterX = False
        End If
        If MouseY < 16 Then
            If YVel < (MouseY - 16) Then
                YVel = YVel + 1
            ElseIf YVel > (MouseY - 16) Then
                YVel = YVel - 1
            End If
            bCenterY = False
        End If
        If MouseY > 463 Then
            If YVel > (MouseY - 463) Then
                YVel = YVel - 1
            ElseIf YVel < (MouseY - 463) Then
                YVel = YVel + 1
            End If
            bCenterY = False
        End If
        If bCenterX Then
            XVel = XVel - Sgn(XVel)
        End If
        If bCenterY Then
            YVel = YVel - Sgn(YVel)
        End If
    End If
    If XVel < -16 Then XVel = -16
    If XVel > 16 Then XVel = 16
    If YVel < -16 Then YVel = -16
    If YVel > 16 Then YVel = 16
    XOff = XOff + XVel
    YOff = YOff + YVel
    If XOff < 0 Then XOff = 0
    If XOff > theMap.MapWidth - theMap.ViewWidth Then
        XOff = theMap.MapWidth - theMap.ViewWidth
    End If
    If YOff < 0 Then YOff = 0
    If YOff > theMap.MapHeight - theMap.ViewHeight Then
        YOff = theMap.MapHeight - theMap.ViewHeight
    End If
End Sub

Sub DrawSpecialFunction(Spec As SpecialFunction)
Attribute DrawSpecialFunction.VB_Description = "Draw a representation of the specified special function taking all scrolling and offsets into account."
    Dim rcSpec As RECT
    Dim hSpecBrush As Long
    Dim hOldBrush As Long
    Dim hOldPen As Long
    Dim hDC As Long
    
    With theMap.MapLayer(EditLayer)
        rcSpec.Left = CLng(Spec.TileLeft) * .TSDef.TileWidth - XOff * .XScrollRate + theMap.ViewLeft
        rcSpec.Top = CLng(Spec.TileTop) * .TSDef.TileHeight - YOff * .YScrollRate + theMap.ViewTop
        rcSpec.Right = CLng(Spec.TileRight + 1) * .TSDef.TileWidth - XOff * .XScrollRate + theMap.ViewLeft
        rcSpec.Bottom = CLng(Spec.TileBottom + 1) * .TSDef.TileHeight - YOff * .YScrollRate + theMap.ViewTop
        If rcSpec.Right < 0 Or rcSpec.Top > 480 Or rcSpec.Left > 640 Or rcSpec.Bottom < 0 Then Exit Sub
        
        hDC = Disp.GetDC
        SetBkMode hDC, Transparent
        hSpecBrush = CreateSolidBrush(vbBlue)
        hOldBrush = SelectObject(hDC, hSpecBrush)
        hOldPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
        SetTextColor hDC, vbYellow
        Rectangle hDC, rcSpec.Left, rcSpec.Top, rcSpec.Right, rcSpec.Bottom
        rcSpec.Top = rcSpec.Top + 1
        rcSpec.Left = rcSpec.Left + 1
        rcSpec.Right = rcSpec.Right - 1
        rcSpec.Bottom = rcSpec.Bottom - 1
        SelectObject hDC, hOldBrush
        DeleteObject hSpecBrush
        SelectObject hDC, hOldPen
        DrawText hDC, Spec.Name, Len(Spec.Name), rcSpec, DT_CENTER + DT_SINGLELINE + DT_VCENTER
        Disp.ReleaseDC
    End With
End Sub

Sub DrawPath(DrawPath As Path, Optional bToMouse As Boolean = False)
Attribute DrawPath.VB_Description = "Draw the specified path taking all scrolling and offsets into account."
    Dim S As Single
    Dim PS As POINTAPI
    Dim PE As POINTAPI
    Dim Dis As Single
    Dim CL As Integer, CT As Integer, CW As Integer, CH As Integer
    Dim I As Integer
    
    With theMap
        CL = .ViewLeft
        CT = .ViewTop
        CW = .ViewWidth
        CH = .ViewHeight
        
        If DrawPath.PointCount > 1 Or ((DrawPath.PointCount = 1) And bToMouse) Then
            With .MapLayer(DrawPath.LayerName)
                PS.X = -XOff * .XScrollRate + DrawPath.PointX(0) + theMap.ViewLeft
                PS.Y = -YOff * .YScrollRate + DrawPath.PointY(0) + theMap.ViewTop
                For I = 1 To DrawPath.PointCount
                    If I < DrawPath.PointCount Then
                        PE.X = -XOff * .XScrollRate + DrawPath.PointX(I) + theMap.ViewLeft
                        PE.Y = -YOff * .YScrollRate + DrawPath.PointY(I) + theMap.ViewTop
                    Else
                        If bToMouse Then
                            PE.X = MouseX
                            PE.Y = MouseY
                            If bPathGrid Then
                                PE.X = ((PE.X - CL + XOff * .XScrollRate) \ .TSDef.TileWidth) * .TSDef.TileWidth - XOff * .XScrollRate + CL
                                PE.Y = ((PE.Y - CT + YOff * .YScrollRate) \ .TSDef.TileHeight) * .TSDef.TileHeight - YOff * .YScrollRate + CT
                            End If
                        Else
                            Exit For
                        End If
                    End If
                    Dis = Sqr((PS.X - PE.X) ^ 2 + (PS.Y - PE.Y) ^ 2)
                    If Dis = 0 Then Dis = 1
                    If Not ((PS.X < CL And PE.X < CL) Or (PS.Y < CT And PE.Y < CT) Or (PS.X > CL + CW And PE.X > CL + CW) Or (PS.Y > CT + CH And PE.Y > CT + CH)) Then
                        For S = FlashTimer / Dis To 1 Step 16 / Dis
                            Disp.DrawTile MouseTS, 3, PS.X * (1 - S) + PE.X * S - 8, PS.Y * (1 - S) + PE.Y * S - 8, CL, CT, CW, CH, True
                        Next
                    End If
                    PS.X = PE.X
                    PS.Y = PE.Y
                Next I
            End With
        ElseIf DrawPath.PointCount > 0 Then
            With .MapLayer(DrawPath.LayerName)
                Disp.DrawTile MouseTS, 3, -XOff * .XScrollRate + DrawPath.PointX(0) + theMap.ViewLeft - 8, -YOff * .YScrollRate + DrawPath.PointY(0) + theMap.ViewTop - 8, CL, CT, CW, CH, True
            End With
        End If
        If bToMouse Then
            With .MapLayer(DrawPath.LayerName)
                PE.X = MouseX
                PE.Y = MouseY
                If bPathGrid Then
                    PE.X = ((PE.X - CL + XOff * .XScrollRate) \ .TSDef.TileWidth) * .TSDef.TileWidth - XOff * .XScrollRate + CL
                    PE.Y = ((PE.Y - CT + YOff * .YScrollRate) \ .TSDef.TileHeight) * .TSDef.TileHeight - YOff * .YScrollRate + CT
                End If
                Disp.DrawTile MouseTS, 3, PE.X - 8, PE.Y - 8, CL, CT, CW, CH, True
            End With
        End If
    End With
End Sub

Sub DrawScreen()
Attribute DrawScreen.VB_Description = "Draw the map, special functions, paths and interactive editor text according to the current state of the editor."
    Dim Col As Integer, Row As Integer
    Dim I As Integer, J As Integer
    Static MatchBlk As Variant
    Static PCol As Integer, PRow As Integer
    Dim hDC As Long
    Dim TextSize As Size
    Dim TextRect As RECT
    Dim Sf As New StdFont
    Dim IFnt As IFont
    Dim Pn As Long
    Dim Op As Long
    Dim ErrStr As String
    
    Sf.Name = "MS Sans Serif"
    Sf.Size = 8
    Sf.Bold = True
    Set IFnt = Sf
    
    On Error GoTo DrawErr
    
    hDC = 0
    
    FlashTimer = (FlashTimer + 1) Mod 16
    
    With theMap
        For I = 0 To .LayerCount - 1
            With .MapLayer(I)
                If theMap.bActive Then
                    If Not .IsLoaded Then
                        On Error Resume Next
                        Disp.Close
                        On Error GoTo 0
                        MsgBox "Layer " & CStr(I) & " is not loaded; Can't draw"
                        Exit Sub
                    End If
                    .AdvanceTileAnims
                End If
                theMap.bActive = True
                Disp.DrawMap .Tileset, .Data, CLng(XOff * .XScrollRate), CLng(YOff * .YScrollRate), theMap.ViewLeft, theMap.ViewTop, theMap.ViewWidth, theMap.ViewHeight, .Transparent
                If Not theMap.bActive Then Exit Sub
                .DrawLayerSprites XOff, YOff
                If bShowAllPaths Then
                    For J = 0 To theMap.PathCount - 1
                        If theMap.Paths(J).LayerName = .Name Then
                            DrawPath theMap.Paths(J)
                        End If
                    Next
                End If
                If I = EditLayer And (EditMode = 11 Or EditMode = 12) Then
                    For J = 0 To theMap.SpecialCount - 1
                        If theMap.Specials(J).LayerIndex = I Then
                            DrawSpecialFunction theMap.Specials(J)
                        End If
                    Next
                End If
            End With
        Next
    End With
    With theMap.MapLayer(EditLayer)
        If EditMode = 9 Or EditMode = 11 Then
            hDC = Disp.GetDC
            SelectObject hDC, IFnt.hFont
            SetBkColor hDC, 0
            SetTextColor hDC, vbWhite
            Pn = CreatePen(PS_SOLID, 1, &HFF8000)
            Op = SelectObject(hDC, Pn)
            With TextRect
                .Left = 0
                .Top = 87
                .Right = 320
                .Bottom = 100
            End With
            If EditMode = 9 Then
                DrawText hDC, "Enter a name for the new path:", 30, TextRect, DT_VCENTER
            Else
                DrawText hDC, "Enter a name for the new special function:", 42, TextRect, DT_VCENTER
            End If
            GetTextExtentPoint32 hDC, ReadString, Len(ReadString), TextSize
            With TextRect
                .Left = 0
                .Top = 100
                .Right = 320
                .Bottom = .Top + 15
            End With
            DrawText hDC, ReadString, Len(ReadString), TextRect, DT_VCENTER
            GetTextExtentPoint32 hDC, Left$(ReadString, ReadPos), ReadPos, TextSize
            Rectangle hDC, TextRect.Left + TextSize.cx, TextRect.Top, TextRect.Left + TextSize.cx + 1, TextRect.Bottom
            SelectObject hDC, Op
            Disp.ReleaseDC
            hDC = 0
            DeleteObject Pn
        ElseIf EditMode = 10 And Not (NewPath Is Nothing) Then
            DrawPath NewPath, True
        ElseIf EditMode = 12 And Not (NewSpec Is Nothing) Then
            If NewSpec.TileRight < 0 Then
                theMap.ScreenToTile XOff, YOff, MouseX, MouseY, EditLayer, Col, Row
                NewSpec.TileLeft = Col
                NewSpec.TileTop = Row
                NewSpec.TileRight = Col
                NewSpec.TileBottom = Row
                DrawSpecialFunction NewSpec
                NewSpec.TileRight = -1
            Else
                theMap.ScreenToTile XOff, YOff, MouseX, MouseY, EditLayer, Col, Row
                NewSpec.TileRight = Col
                NewSpec.TileBottom = Row
                If NewSpec.TileRight < NewSpec.TileLeft Then NewSpec.TileLeft = NewSpec.TileRight
                If NewSpec.TileBottom < NewSpec.TileTop Then NewSpec.TileTop = NewSpec.TileBottom
                DrawSpecialFunction NewSpec
            End If
        End If
    End With
    If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, Col, Row) Then
        Disp.ForeColor = vbWhite
        Disp.DrawText CStr(Col) & ", " & CStr(Row), 0, 0
        With theMap.MapLayer(EditLayer)
            If EditMode = 0 Then
                Disp.DrawTile .Tileset, CurTile, theMap.ViewLeft + CLng(Col) * .TSDef.TileWidth - XOff * theMap.MapLayer(EditLayer).XScrollRate, theMap.ViewTop + CLng(Row) * .TSDef.TileHeight - YOff * theMap.MapLayer(EditLayer).YScrollRate, theMap.ViewLeft, theMap.ViewTop, theMap.ViewWidth, theMap.ViewHeight, False
            ElseIf EditMode = 4 Then
                If theMap.ScreenToTile(XOff, YOff, MouseX, MouseY, EditLayer, Col, Row) Then
                    For J = Row To Row + UBound(CopyData, 2)
                        For I = Col To Col + UBound(CopyData, 1)
                            Disp.DrawTile .Tileset, CopyData(I - Col, J - Row), theMap.ViewLeft + CLng(I) * .TSDef.TileWidth - XOff * theMap.MapLayer(EditLayer).XScrollRate, theMap.ViewTop + CLng(J) * .TSDef.TileHeight - YOff * theMap.MapLayer(EditLayer).YScrollRate, theMap.ViewLeft, theMap.ViewTop, theMap.ViewWidth, theMap.ViewHeight, False
                        Next
                    Next
                End If
            ElseIf EditMode = 5 Or EditMode = 6 Then
                If theMap.ScreenToTile(XOff, YOff, MouseX - .Tileset.TileWidth / 2, MouseY - .Tileset.TileHeight / 2, EditLayer, Col, Row) Then
                    If PCol <> Col Or PRow <> Row Then
                        MatchBlk = Match2x2(Col, Row, EditMode = 6)
                        PCol = Col
                        PRow = Row
                    End If
                    For J = Row - 1 To Row + 2
                        For I = Col - 1 To Col + 2
                            If I >= 0 And I < theMap.MapLayer(EditLayer).Columns And J >= 0 And J < theMap.MapLayer(EditLayer).Rows Then
                                If MatchBlk(I - Col + 1, J - Row + 1) <> theMap.MapLayer(EditLayer).Data.TileValue(I, J) Then
                                    Disp.DrawTile .Tileset, MatchBlk(I - Col + 1, J - Row + 1), theMap.ViewLeft + CLng(I) * .TSDef.TileWidth - XOff * theMap.MapLayer(EditLayer).XScrollRate, theMap.ViewTop + CLng(J) * .TSDef.TileHeight - YOff * theMap.MapLayer(EditLayer).YScrollRate, theMap.ViewLeft, theMap.ViewTop, theMap.ViewWidth, theMap.ViewHeight, False
                                End If
                            End If
                        Next
                    Next
                End If
            End If
        End With
    End If
    
    If Len(DisplayMessage) > 0 Then
        Disp.DrawText DisplayMessage, 0, 200, 640
        Disp.DrawText "Press enter to clear this message", 0, 216, 640
    End If
    
    Exit Sub
    
DrawErr:
    ErrStr = Err.Description
    If hDC <> 0 Then
        Disp.ReleaseDC
    End If
    bQuit = True
    Disp.Close
    Set Disp = Nothing
    Set CurDisp = Nothing
    MsgBox ErrStr, vbExclamation
End Sub

Function Match2x2(ByVal LCol As Integer, ByVal TRow As Integer, Optional bErase As Boolean = False) As Variant
Attribute Match2x2.VB_Description = "Return a 4x4 array of tiles that matches a 2x2 MatchDef block to a piece of the current layer."
    Dim R As Integer, C As Integer
    Dim InR As Integer, InC As Integer
    Dim TileVals(0 To 2, 0 To 2) As Byte
    Dim RetVal As Variant
    Dim MaxCol As Integer
    Dim MaxRow As Integer
    
    MaxCol = theMap.MapLayer(EditLayer).Columns - 1
    MaxRow = theMap.MapLayer(EditLayer).Rows - 1
    
    ReDim RetVal(0 To 3, 0 To 3) As Byte
    
    With theMap.MapLayer(EditLayer).Data
        For C = LCol - 1 To LCol + 2
            For R = TRow - 1 To TRow + 2
                For InC = C - 1 To C + 1
                    For InR = R - 1 To R + 1
                        If InR < TRow Or InR > TRow + 1 Or InC < LCol Or InC > LCol + 1 Then
                            If InR >= 0 And InR <= MaxRow And InC >= 0 And InC <= MaxCol Then
                                TileVals(InC - C + 1, InR - R + 1) = .TileValue(InC, InR)
                            Else
                                TileVals(InC - C + 1, InR - R + 1) = 0
                            End If
                        Else
                            If bErase Then
                                TileVals(InC - C + 1, InR - R + 1) = 0
                            Else
                                TileVals(InC - C + 1, InR - R + 1) = UseMatchDef.TileMatches.MatchGroup(6).GetMember(0)
                            End If
                        End If
                    Next
                Next
                If C >= 0 And R >= 0 And C <= MaxCol And R <= MaxRow Then
                    If UseMatchDef.AllTiles.IsMember(TileVals(1, 1)) Then
                        RetVal(C - LCol + 1, R - TRow + 1) = _
                            UseMatchDef.MatchMiddleTile(TileVals(0, 0), TileVals(1, 0), TileVals(2, 0), _
                                                        TileVals(0, 1), TileVals(2, 1), _
                                                        TileVals(0, 2), TileVals(1, 2), TileVals(2, 2))
                    Else
                        RetVal(C - LCol + 1, R - TRow + 1) = TileVals(1, 1)
                    End If
                End If
            Next
        Next
    End With
    
    Match2x2 = RetVal
    
End Function

Sub PaintMenu()
Attribute PaintMenu.VB_Description = "Draw the menus according to the current state."
    Dim I As Integer
    Dim S As New StdFont
    
    S.Name = "MS Sans Serif"
    S.Size = 8
    S.Bold = True
    Disp.SetFont S
    Set S = Nothing
    
    For I = LBound(MenuItems) To UBound(MenuItems)
        With MenuItems(I)
            If MouseX > .rcBound.Left And MouseX < .rcBound.Right And _
               MouseY > .rcBound.Top And MouseY < .rcBound.Bottom Then
                Disp.ForeColor = vbRed
                If I = 2 Then MenuMode = 1
                If I = 3 Then MenuMode = 3
                If I = 4 Then MenuMode = 2
            Else
                Disp.ForeColor = vbWhite
            End If
            Disp.DrawText .Caption, .rcBound.Left, .rcBound.Top, .rcBound.Right - .rcBound.Left + 1
        End With
    Next
    If MenuMode = 1 Then
        For I = 0 To Prj.GetApplicableTileMatchCount(theMap.MapLayer(EditLayer).TSDef) - 1
            With MenuItems(2).rcBound
                If MouseY >= .Bottom + I * 20 And MouseY < .Bottom + I * 20 + 20 Then
                    Disp.ForeColor = vbRed
                Else
                    Disp.ForeColor = vbWhite
                End If
                Disp.DrawText Prj.GetApplicableTileMatch(theMap.MapLayer(EditLayer).TSDef, I).Name, .Left, .Bottom + I * 20
                If MouseY < .Bottom + I * 20 + 10 Then Disp.ForeColor = vbWhite
                Disp.DrawText "Eraser", .Left, .Bottom + I * 20 + 10
            End With
        Next
        With MenuItems(2).rcBound
            If MouseX < .Left Or MouseX > .Right Or MouseY > .Bottom + I * 20 Then MenuMode = 0
        End With
    ElseIf MenuMode = 2 Then
        For I = LBound(PathMenu) To UBound(PathMenu)
            With PathMenu(I)
                If MouseY >= .rcBound.Top And MouseY <= .rcBound.Bottom Then
                    Disp.ForeColor = vbRed
                Else
                    Disp.ForeColor = vbWhite
                End If
                Disp.DrawText .Caption, .rcBound.Left, .rcBound.Top
                If MouseX < .rcBound.Left Or MouseX > .rcBound.Right Or MouseY > .rcBound.Bottom + 30 Then MenuMode = 0
            End With
        Next
    ElseIf MenuMode = 3 Then
        For I = LBound(MazeMenu) To UBound(MazeMenu)
            With MazeMenu(I)
                If MouseY >= .rcBound.Top And MouseY <= .rcBound.Bottom Then
                    Disp.ForeColor = vbRed
                Else
                    Disp.ForeColor = vbWhite
                End If
                Disp.DrawText .Caption, .rcBound.Left, .rcBound.Top
                If MouseX < .rcBound.Left Or MouseX > .rcBound.Right Or MouseY > .rcBound.Bottom + 20 Then MenuMode = 0
            End With
        Next
    Else
        With theMap.MapLayer(EditLayer)
            Disp.CopyRect .Tileset, 0, 0, Disp.HimetricToPixelX(.TSDef.Image.Width), Disp.HimetricToPixelY(.TSDef.Image.Height), 0, -TSOffset + MenuHeight, 0, MenuHeight, 640, 480 - MenuHeight, False
            If (MouseX \ .TSDef.TileWidth) * .TSDef.TileWidth + .TSDef.TileWidth <= Disp.HimetricToPixelX(.TSDef.Image.Width) And _
                ((MouseY - MenuHeight + TSOffset) \ .TSDef.TileHeight) * .TSDef.TileHeight >= 0 And ((MouseY - MenuHeight + TSOffset) \ .TSDef.TileHeight) * .TSDef.TileHeight + .TSDef.TileHeight <= Disp.HimetricToPixelY(.TSDef.Image.Height) Then
                Disp.PaintPicture ExtractTile(.TSDef.Image, (MouseX \ .TSDef.TileWidth) * .TSDef.TileWidth, ((MouseY - MenuHeight + TSOffset) \ .TSDef.TileHeight) * .TSDef.TileHeight, .TSDef.TileWidth, .TSDef.TileHeight, True), _
                    (MouseX \ .TSDef.TileWidth) * .TSDef.TileWidth, ((MouseY - MenuHeight + TSOffset) \ .TSDef.TileHeight) * .TSDef.TileHeight + MenuHeight - TSOffset
            End If
        End With
    End If
    
End Sub

Sub ActivateMenu(ByVal X As Integer, ByVal Y As Integer)
Attribute ActivateMenu.VB_Description = "Respond to a click on the menu screen"
    Dim I As Integer
    
    On Error GoTo ActivateErr

    With theMap.MapLayer(EditLayer)
        If MenuMode = 1 Then
            I = (MouseY - MenuItems(2).rcBound.Bottom) \ 10
            If I \ 2 >= 0 And I \ 2 < Prj.GetApplicableTileMatchCount(theMap.MapLayer(EditLayer).TSDef) Then
                Set UseMatchDef = Prj.GetApplicableTileMatch(theMap.MapLayer(EditLayer).TSDef, I \ 2)
                If I Mod 2 > 0 Then
                    EditMode = 6
                Else
                    EditMode = 5
                End If
            End If
        ElseIf MenuMode = 2 Then
            I = (MouseY - MenuItems(4).rcBound.Bottom) \ 10
            Select Case I
            Case 0
                EditMode = 9
            Case 1
                bShowAllPaths = Not bShowAllPaths
                If bShowAllPaths Then
                    PathMenu(1).Caption = "Hide All"
                Else
                    PathMenu(1).Caption = "Show All"
                End If
                EditMode = 0
            Case 2
                theMap.InitSprites
                Prj.GamePlayer.InitPlayerSprite
                theMap.bDisablePlayerEdit = (Val(GetSetting("GameDev", "Options", "DisablePlayerEdit", "0")) <> 0)
                EditMode = 0
            Case 3 ' Toggle Path Grid
                bPathGrid = Not bPathGrid
                EditMode = 0
            End Select
            MenuMode = 0
        ElseIf MenuMode = 3 Then
            I = (MouseY - MenuItems(3).rcBound.Bottom) \ 10
            Select Case I
            Case 0
                MzWS = 6
                MzPS = 1
                EditMode = 7
            Case 1
                MzWS = 8
                MzPS = 2
                EditMode = 7
            End Select
            MenuMode = 0
        Else
            If (Y - MenuHeight + TSOffset) >= 0 And (Y - MenuHeight + TSOffset) < Disp.HimetricToPixelY(.TSDef.Image.Height) Then
                CurTile = X \ .TSDef.TileWidth + ((Y - MenuHeight + TSOffset) \ .TSDef.TileHeight) * .Tileset.TileColumns
                EditMode = 0
            Else
                For I = LBound(MenuItems) To UBound(MenuItems)
                    With MenuItems(I).rcBound
                        If X > .Left And X < .Right And Y > .Top And Y < .Bottom Then
                            Exit For
                        End If
                    End With
                Next I
                Select Case I
                    Case 0
                        bQuit = True
                    Case 1
                        EditMode = 2
                    Case 5
                        EditMode = 11
                End Select
            End If
        End If
    End With
    If (EditMode < 5 Or EditMode > 8) And GeneratingMaze = False Then Set UseMatchDef = Nothing
    Exit Sub
     
ActivateErr:
     Disp.Close
     MsgBox Err.Description
End Sub

