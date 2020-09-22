VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The object used to handle playing the game, and the player sprite's control of the game.  Owned by project."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Player.cls - Game Player and Player Sprite Handling Class
'
'======================================================================

Option Explicit

' Persistent
Public ScrollMarginX As Integer
Attribute ScrollMarginX.VB_VarDescription = "Get or set how close the edge of the player sprite can get to the edge of the map window before scrolling."
Public ScrollMarginY As Integer
Attribute ScrollMarginY.VB_VarDescription = "Get or set how close the edge of the player sprite can get to the edge of the map window before scrolling."
Public InvBarMargin As Integer
Attribute InvBarMargin.VB_VarDescription = "Set the margin between all inventory item displays and their respective quantity display."
Public StartMapName As String
Attribute StartMapName.VB_VarDescription = "Set the name of the map where play initially starts."
Private KeyCfg(0 To 7) As Integer
Public bEnableJoystick As Boolean
Attribute bEnableJoystick.VB_VarDescription = "Allow joystick control of the player."

' Not persistent
Public MapScrollX As Long
Attribute MapScrollX.VB_VarDescription = "Get or set the current horizontal scroll position of the map."
Public MapScrollY As Long
Attribute MapScrollY.VB_VarDescription = "Get or set the current vertical scroll position of the map."
Public rMap As Map
Attribute rMap.VB_VarDescription = "Reference to the map object currently being displayed by the game player."
Public PlayerSprite As Sprite
Attribute PlayerSprite.VB_VarDescription = "After InitPlayerSprite is called, this refers to the player sprite object."
Private JSInfo As JOYINFOEX
Public JSValid As Boolean
Attribute JSValid.VB_VarDescription = "Return True if joystick input was valid after the last PollJoystick call."
Public WithEvents Disp As BMDXDisplay
Attribute Disp.VB_VarHelpID = -1
Public WithEvents Touch As MapInteract
Attribute Touch.VB_VarHelpID = -1
Public bQuit As Boolean
Attribute bQuit.VB_VarDescription = "Set to True to end main Play loop."
Public InterestingTiles As TileGroup
Attribute InterestingTiles.VB_VarDescription = "Tiles which should be considered for raising an event when touched. (Optimization for MapInteract.TouchTest)"
Public FirstTouchTiles As TileGroup
Attribute FirstTouchTiles.VB_VarDescription = "A TileGroup to mask in all tiles which are only interested when initially touched.  (optimization for MapInteract.TouchTest.)"
Public KeyActions As Integer
Attribute KeyActions.VB_VarDescription = "A bitmask of eActionBits flags indicating the current state of the keyboard."
Public JoyActions As Integer
Attribute JoyActions.VB_VarDescription = "A bitmask of eActionBits flags indicating the current state of the joystick."
Public CtlActions As Integer
Attribute CtlActions.VB_VarDescription = "A bitmask of eActionBits flags indicating the current state of the input controllers ORed together."

Private bPause As Boolean

Public Enum QuantityDisplay
    QD_NONE = 0
    QD_HORZBARABOVE
    QD_HORZBARBELOW
    QD_HORZBARRIGHT
    QD_VERTBARABOVE
    QD_VERTBARRIGHT
    QD_VERTBARLEFT
    QD_TEXTABOVE
    QD_TEXTBELOW
    QD_TEXTLEFT
    QD_TEXTRIGHT
    QD_REPEATICONRIGHT
    QD_REPEATICONABOVE
    QD_REPEATICONBELOW
    QD_INCREASEICONINDEX
    QD_ICONINDEXREPEATRIGHT
    QD_ICONINDEXREPEATABOVE
    QD_ICONINDEXREPEATBELOW
End Enum

Private Type InvItem
    Name As String
    IconTilesetName As String
    IconTilesetIndex As Integer
    IconTileIndex As Integer
    QuantityOwned As Integer
    MaxQuantity As Integer
    QuantityDisplayType As QuantityDisplay
    IconCountPerRepeat As Integer
    BarColor As Long
    BarBackgroundColor As Long
    BarOutlineColor As Long
    BarLength As Integer
    BarThickness As Integer
    DisplayX As Integer
    DisplayY As Integer
End Type

Enum eActionBits
    ACTION_UP = 1
    ACTION_LEFT = 2
    ACTION_RIGHT = 4
    ACTION_DOWN = 8
    ACTION_BUTTON1 = 16
    ACTION_BUTTON2 = 32
    ACTION_BUTTON3 = 64
    ACTION_BUTTON4 = 128
End Enum

Private Inventory() As InvItem
Private InvCount As Integer

Event OnTileInteraction(ByVal TileX As Long, ByVal TileY As Long, ByVal TileValue As Integer, bNewTouch As Boolean, InteractionIndex As Integer)
Attribute OnTileInteraction.VB_Description = "Fired when a map tile interaction raises an event."
Event OnSpecialFunction(SpecialObj As SpecialFunction)
Attribute OnSpecialFunction.VB_Description = "Fired when the player touches a special function that raises an event."
Event OnControllerMove(ByVal OldActions As Integer, ByVal NewActions As Integer)
Attribute OnControllerMove.VB_Description = "Fired when an input controller moves (CtlActions changes)."
Event OnBeforeMoveSprites()
Attribute OnBeforeMoveSprites.VB_Description = "Fired right before all the sprite movements are processed."
Event OnAfterMoveSprites()
Attribute OnAfterMoveSprites.VB_Description = "Fired right after all the sprite movements are processed."
Event OnSpritesCollide(ByVal Name As String, ByVal ClsASprIdx As Integer, ByVal ClsBSprIdx As Integer, ByVal CollDefIdx As Integer)
Attribute OnSpritesCollide.VB_Description = "Fired when sprites collide."
Event OnPlayInit()
Attribute OnPlayInit.VB_Description = "Fired after initialization is complete, right before Play loop starts."
Event OnNewSprite(NewSprite As Sprite)
Attribute OnNewSprite.VB_Description = "Fired when a new sprite is added to a layer."
Event OnPlayComplete()
Attribute OnPlayComplete.VB_Description = "Fired after all the clean up is done after exiting the game player Play loop."

Public Property Get KeyConfig(ByVal Idx As Integer) As Integer
Attribute KeyConfig.VB_Description = "Get or set the virtual key used for an action: Up, Left, Right, Down, Buttons 1-4."
    KeyConfig = KeyCfg(Idx)
End Property

Public Property Let KeyConfig(ByVal Idx As Integer, ByVal NewVal As Integer)
    KeyCfg(Idx) = NewVal
End Property

Public Property Get InventoryCount() As Integer
Attribute InventoryCount.VB_Description = "Return the number of defined inventory items."
    InventoryCount = InvCount
End Property

Public Function AddInventoryItem(ByVal QuantityDisplayType As QuantityDisplay, ByVal DisplayX As Integer, ByVal DisplayY As Integer) As Integer
Attribute AddInventoryItem.VB_Description = "Add a new inventory item definition."
    If InvCount = 0 Then
        ReDim Inventory(0) As InvItem
    Else
        ReDim Preserve Inventory(0 To InvCount) As InvItem
    End If
    With Inventory(InvCount)
        .QuantityDisplayType = QuantityDisplayType
        .DisplayX = DisplayX
        .DisplayY = DisplayY
    End With
    AddInventoryItem = InvCount
    InvCount = InvCount + 1
End Function

Public Sub RemoveInventoryItem(ByVal Index As Integer)
Attribute RemoveInventoryItem.VB_Description = "Remove an inventory item definition."
    Dim I As Integer
    
    For I = Index To InvCount - 2
        Inventory(I) = Inventory(I + 1)
    Next
    
    InvCount = InvCount - 1
    If InvCount > 0 Then
        ReDim Preserve Inventory(0 To InvCount - 1)
    Else
        Erase Inventory
    End If
        
End Sub

Public Property Let InventoryItemName(ByVal Index As Integer, ByVal NewVal As String)
Attribute InventoryItemName.VB_Description = "Get or set the display name of an item in the inventory."
   Inventory(Index).Name = NewVal
End Property

Public Property Get InventoryItemName(ByVal Index As Integer) As String
    InventoryItemName = Inventory(Index).Name
End Property

Public Sub SetInventoryTile(ByVal Index As Integer, ByVal TileSetKey As Variant, ByVal TileIndex As Integer)
Attribute SetInventoryTile.VB_Description = "Set the tileset (by name or index) and tile index used to display a particular inventory item."
    Inventory(Index).IconTilesetName = ""
    On Error Resume Next
    If VarType(TileSetKey) = vbString Then
        Inventory(Index).IconTilesetIndex = Prj.TileSetDef(TileSetKey).GetIndex
        If Err.Number = 0 Then
            Inventory(Index).IconTilesetName = TileSetKey
        End If
    Else
        Inventory(Index).IconTilesetName = Prj.TileSetDef(TileSetKey).Name
        Inventory(Index).IconTilesetIndex = TileSetKey
    End If
    Inventory(Index).IconTileIndex = TileIndex
End Sub

Public Sub ReIndexTilesetRefs()
Attribute ReIndexTilesetRefs.VB_Description = "Find the indexes of the tilesets named in the inventory to prepare for efficient inventory drawing."
    Dim I As Integer

    For I = 0 To InvCount - 1
        If Inventory(I).IconTilesetName = "" Then
            Inventory(I).IconTilesetIndex = -1
        Else
            Inventory(I).IconTilesetIndex = Prj.TileSetDef(Inventory(I).IconTilesetName).GetIndex
        End If
    Next
End Sub

Public Property Get InvIconTilesetIdx(ByVal Index As Integer) As Integer
Attribute InvIconTilesetIdx.VB_Description = "Return the ordinal index in the project of the TileSetDef containing the graphics for this inventory item."
    InvIconTilesetIdx = Inventory(Index).IconTilesetIndex
End Property

Public Property Get InvIconTileIdx(ByVal Index As Integer) As Integer
Attribute InvIconTileIdx.VB_Description = "Return the index of the tile used to represent an inventory item."
    InvIconTileIdx = Inventory(Index).IconTileIndex
End Property

Public Property Let InvQuantityOwned(ByVal Index As Integer, ByVal NewVal As Integer)
Attribute InvQuantityOwned.VB_Description = "Get or set how many of an inventory item the player currently owns."
    Inventory(Index).QuantityOwned = NewVal
End Property

Public Property Get InvQuantityOwned(ByVal Index As Integer) As Integer
    InvQuantityOwned = Inventory(Index).QuantityOwned
End Property

Public Property Let InvMaxQuantity(ByVal Index As Integer, ByVal NewVal As Integer)
Attribute InvMaxQuantity.VB_Description = "Get or set the maximum number of an item that the player can own."
    Inventory(Index).MaxQuantity = NewVal
End Property

Public Property Get InvMaxQuantity(ByVal Index As Integer) As Integer
    InvMaxQuantity = Inventory(Index).MaxQuantity
End Property

Public Property Get InvQuantityDisplayType(ByVal Index As Integer) As QuantityDisplay
Attribute InvQuantityDisplayType.VB_Description = "Get or set how an inventory item should be displayed."
    InvQuantityDisplayType = Inventory(Index).QuantityDisplayType
End Property

Public Property Let InvQuantityDisplayType(ByVal Index As Integer, ByVal NewVal As QuantityDisplay)
    Inventory(Index).QuantityDisplayType = NewVal
End Property

Public Property Let InvIconCountPerRepeat(ByVal Index As Integer, ByVal NewVal As Integer)
Attribute InvIconCountPerRepeat.VB_Description = "How many images of this inventory item are there to represent owning 1, 2, 3 (etc) of this item."
    Inventory(Index).IconCountPerRepeat = NewVal
End Property

Public Property Get InvIconCountPerRepeat(ByVal Index As Integer) As Integer
    InvIconCountPerRepeat = Inventory(Index).IconCountPerRepeat
End Property

Public Sub InvSetBarInfo(ByVal Index As Integer, ByVal BarColor As Long, ByVal BarThickness As Integer, _
                         ByVal BarLength As Integer, Optional ByVal BackgroundColor As Long = -1, _
                         Optional ByVal OutlineColor As Long = -1)
Attribute InvSetBarInfo.VB_Description = "Set all aspects of how a bar is displayed for an inventory item."
    With Inventory(Index)
        .BarColor = BarColor
        .BarThickness = BarThickness
        .BarLength = BarLength
        .BarBackgroundColor = BackgroundColor
        .BarOutlineColor = OutlineColor
    End With
End Sub

Public Property Get InvBarColor(ByVal Index As Integer) As Long
Attribute InvBarColor.VB_Description = "Return the color for a bar on an inventory item display."
    InvBarColor = Inventory(Index).BarColor
End Property

Public Property Get InvBarThickness(ByVal Index As Integer) As Integer
Attribute InvBarThickness.VB_Description = "Return the thickness of a bar in the inventory item display."
    InvBarThickness = Inventory(Index).BarThickness
End Property

Public Property Get InvBarLength(ByVal Index As Integer) As Integer
Attribute InvBarLength.VB_Description = "Return the total length of an inventory item's display bar in pixels."
    InvBarLength = Inventory(Index).BarLength
End Property

Public Property Get InvBarBackgroundColor(ByVal Index As Integer) As Long
Attribute InvBarBackgroundColor.VB_Description = "Return the background color for a bar on an inventory item display."
    InvBarBackgroundColor = Inventory(Index).BarBackgroundColor
End Property

Public Property Get InvBarOutlineColor(ByVal Index As Integer) As Long
Attribute InvBarOutlineColor.VB_Description = "Return the outline color for a bar on an inventory item display."
    InvBarOutlineColor = Inventory(Index).BarOutlineColor
End Property

Public Sub InvMove(ByVal Index As Integer, ByVal DisplayX As Integer, ByVal DisplayY As Integer)
Attribute InvMove.VB_Description = "Change the X and Y coordinates of an inventory item's display."
    With Inventory(Index)
        .DisplayX = DisplayX
        .DisplayY = DisplayY
    End With
End Sub

Public Property Get InvDisplayX(ByVal Index As Integer) As Integer
Attribute InvDisplayX.VB_Description = "Return the X coordinate for the display of a particular inventory item."
    InvDisplayX = Inventory(Index).DisplayX
End Property

Public Property Get InvDisplayY(ByVal Index As Integer) As Integer
Attribute InvDisplayY.VB_Description = "Return the Y coordinate for the display of a particular inventory item."
    InvDisplayY = Inventory(Index).DisplayY
End Property

' If hDC <> NULL, draw all hDC content only
' if hDC = NULL, draw tiles only
Public Sub DrawInvItem(ByVal Index As Integer, Optional ByVal hDC As Long = 0)
Attribute DrawInvItem.VB_Description = "Draw a single inventory item onto the display. hDC=NULL: draw only tiles. hDC <> NULL: draw only lines/text."
    On Error GoTo DrawInvError
    
    Dim CurRight As Integer
    Dim CurBottom As Integer
    Dim sSize As Size
    Dim hBarBrush As Long
    Dim hBarPen As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim rcBarRect As RECT
    Dim rcOutlineRect As RECT
    Dim TmpStr As String
    Dim Idx As Integer
    Dim DrawIdx As Integer
    Dim ErrDesc As String

    With Inventory(Index)
        If .QuantityDisplayType = QD_NONE Then Exit Sub

        If .IconTilesetName <> "" Then
            If hDC = 0 Then
                With Prj.TileSetDef(.IconTilesetIndex)
                    If .LoadedInstance Is Nothing Then
                        If Not .IsLoaded Then
                            .Load
                        End If
                        Set .LoadedInstance = Disp.CreateTileSet(.Image, .TileWidth, .TileHeight)
                    End If
                End With
                Select Case .QuantityDisplayType
                Case QD_INCREASEICONINDEX, QD_ICONINDEXREPEATRIGHT, QD_ICONINDEXREPEATABOVE, QD_ICONINDEXREPEATBELOW, QD_REPEATICONABOVE, QD_REPEATICONRIGHT, QD_REPEATICONBELOW
                    'Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, 0, .DisplayX, .DisplayY, 0, 0, 640, 480, False
                Case Else
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, .IconTileIndex, .DisplayX, .DisplayY, 0, 0, 640, 480, False
                End Select
            End If
            CurRight = .DisplayX + Prj.TileSetDef(.IconTilesetIndex).TileWidth
            CurBottom = .DisplayY + Prj.TileSetDef(.IconTilesetIndex).TileHeight
        Else
            If hDC <> 0 Then
                Disp.DrawText .Name, .DisplayX, .DisplayY
                GetTextExtentPoint32 hDC, .Name, Len(.Name), sSize
                Select Case .QuantityDisplayType
                Case QD_REPEATICONRIGHT, QD_REPEATICONABOVE, QD_REPEATICONBELOW, QD_INCREASEICONINDEX, _
                     QD_ICONINDEXREPEATRIGHT, QD_ICONINDEXREPEATABOVE, QD_ICONINDEXREPEATBELOW
                     Exit Sub
                End Select
            End If
            CurRight = .DisplayX + sSize.cx - 1
            CurBottom = .DisplayY + sSize.cy - 1
        End If
        Select Case .QuantityDisplayType
        Case QD_HORZBARABOVE, QD_HORZBARBELOW, QD_HORZBARRIGHT, QD_VERTBARABOVE, QD_VERTBARRIGHT, QD_VERTBARLEFT
            If hDC <> 0 Then
                Select Case .QuantityDisplayType
                Case QD_HORZBARABOVE
                    rcOutlineRect.Left = .DisplayX
                    rcOutlineRect.Top = .DisplayY - InvBarMargin - .BarThickness
                    rcOutlineRect.Right = .DisplayX + .BarLength
                    rcOutlineRect.Bottom = .DisplayY - InvBarMargin
                    rcBarRect = rcOutlineRect
                    rcBarRect.Right = rcOutlineRect.Left + (.QuantityOwned * .BarLength) / .MaxQuantity
                Case QD_HORZBARBELOW
                    rcOutlineRect.Left = .DisplayX
                    rcOutlineRect.Top = CurBottom + InvBarMargin
                    rcOutlineRect.Right = .DisplayX + .BarLength
                    rcOutlineRect.Bottom = CurBottom + InvBarMargin + .BarThickness
                    rcBarRect = rcOutlineRect
                    rcBarRect.Right = rcOutlineRect.Left + (.QuantityOwned * .BarLength) / .MaxQuantity
                Case QD_HORZBARRIGHT
                    rcOutlineRect.Left = CurRight + InvBarMargin
                    rcOutlineRect.Top = (.DisplayY + CurBottom - .BarThickness) / 2
                    rcOutlineRect.Right = CurRight + InvBarMargin + .BarLength
                    rcOutlineRect.Bottom = (.DisplayY + CurBottom + .BarThickness) / 2
                    rcBarRect = rcOutlineRect
                    rcBarRect.Right = rcOutlineRect.Left + (.QuantityOwned * .BarLength) / .MaxQuantity
                Case QD_VERTBARABOVE
                    rcOutlineRect.Left = (.DisplayX + CurRight - .BarThickness) / 2
                    rcOutlineRect.Top = .DisplayY - InvBarMargin - .BarLength
                    rcOutlineRect.Right = (.DisplayX + CurRight + .BarThickness) / 2
                    rcOutlineRect.Bottom = .DisplayY - InvBarMargin
                    rcBarRect = rcOutlineRect
                    rcBarRect.Top = .DisplayY - InvBarMargin - (.QuantityOwned * .BarLength) / .MaxQuantity
                Case QD_VERTBARRIGHT
                    rcOutlineRect.Left = CurRight + InvBarMargin
                    rcOutlineRect.Top = CurBottom - .BarLength
                    rcOutlineRect.Right = CurRight + InvBarMargin + .BarThickness
                    rcOutlineRect.Bottom = CurBottom
                    rcBarRect = rcOutlineRect
                    rcBarRect.Top = CurBottom - (.QuantityOwned * .BarLength) / .MaxQuantity
                Case QD_VERTBARLEFT
                    rcOutlineRect.Left = .DisplayX - InvBarMargin - .BarThickness
                    rcOutlineRect.Top = CurBottom - .BarLength
                    rcOutlineRect.Right = .DisplayX - InvBarMargin
                    rcOutlineRect.Bottom = CurBottom
                    rcBarRect = rcOutlineRect
                    rcBarRect.Top = CurBottom - (.QuantityOwned * .BarLength) / .MaxQuantity
                End Select
                hBarBrush = CreateSolidBrush(.BarBackgroundColor)
                FillRect hDC, rcOutlineRect, hBarBrush
                DeleteObject hBarBrush
                hBarBrush = CreateSolidBrush(.BarColor)
                FillRect hDC, rcBarRect, hBarBrush
                DeleteObject hBarBrush
                If .BarOutlineColor <> -1 Then
                    hBarPen = CreatePen(PS_SOLID, 1, .BarOutlineColor)
                    hOldPen = SelectObject(hDC, hBarPen)
                    hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
                    With rcOutlineRect
                        Rectangle hDC, .Left, .Top, .Right, .Bottom
                    End With
                    SelectObject hDC, hOldPen
                    DeleteObject hBarPen
                End If
            End If
        Case QD_TEXTABOVE
            If hDC <> 0 Then
                TmpStr = CStr(.QuantityOwned) & "  "
                GetTextExtentPoint32 hDC, TmpStr, Len(TmpStr), sSize
                rcBarRect.Left = .DisplayX
                rcBarRect.Top = .DisplayY - InvBarMargin - sSize.cy
                With rcBarRect
                    .Right = .Left + sSize.cx
                    .Bottom = .Top + sSize.cy
                End With
                SetBkColor hDC, 0
                SetTextColor hDC, vbWhite
                DrawText hDC, TmpStr, -1, rcBarRect, 0
            End If
        Case QD_TEXTBELOW
            If hDC <> 0 Then
                TmpStr = CStr(.QuantityOwned) & "  "
                GetTextExtentPoint32 hDC, TmpStr, Len(TmpStr), sSize
                rcBarRect.Left = .DisplayX
                rcBarRect.Top = CurBottom + InvBarMargin
                With rcBarRect
                    .Right = .Left + sSize.cx
                    .Bottom = .Top + sSize.cy
                End With
                SetBkColor hDC, 0
                SetTextColor hDC, vbWhite
                DrawText hDC, TmpStr, -1, rcBarRect, 0
            End If
        Case QD_TEXTLEFT
            If hDC <> 0 Then
                TmpStr = CStr(.QuantityOwned) & "  "
                GetTextExtentPoint32 hDC, TmpStr, Len(TmpStr), sSize
                rcBarRect.Left = .DisplayX - InvBarMargin - sSize.cx
                rcBarRect.Top = (.DisplayY + CurBottom - sSize.cy) / 2
                With rcBarRect
                    .Right = .Left + sSize.cx
                    .Bottom = .Top + sSize.cy
                End With
                SetBkColor hDC, 0
                SetTextColor hDC, vbWhite
                DrawText hDC, TmpStr, -1, rcBarRect, 0
            End If
        Case QD_TEXTRIGHT
            If hDC <> 0 Then
                TmpStr = CStr(.QuantityOwned) & "  "
                GetTextExtentPoint32 hDC, TmpStr, Len(TmpStr), sSize
                rcBarRect.Left = CurRight + InvBarMargin
                rcBarRect.Top = (.DisplayY + CurBottom - sSize.cy) / 2
                With rcBarRect
                    .Right = .Left + sSize.cx
                    .Bottom = .Top + sSize.cy
                End With
                SetBkColor hDC, 0
                SetTextColor hDC, vbWhite
                DrawText hDC, TmpStr, -1, rcBarRect, 0
            End If
        Case QD_REPEATICONRIGHT
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To .MaxQuantity - 1
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, IIf(Idx < .QuantityOwned, .IconTileIndex, 0), _
                        .DisplayX + Prj.TileSetDef(.IconTilesetIndex).TileWidth * Idx, .DisplayY, 0, 0, 640, 480, False
                Next
            End If
        Case QD_REPEATICONABOVE
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To .MaxQuantity - 1
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, IIf(Idx < .QuantityOwned, .IconTileIndex, 0), _
                        .DisplayX, .DisplayY - Prj.TileSetDef(.IconTilesetIndex).TileHeight * Idx, 0, 0, 640, 480, False
                Next
            End If
        Case QD_REPEATICONBELOW
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To .MaxQuantity - 1
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, IIf(Idx < .QuantityOwned, .IconTileIndex, 0), _
                        .DisplayX, .DisplayY + Prj.TileSetDef(.IconTilesetIndex).TileHeight * Idx, 0, 0, 640, 480, False
                Next
            End If
        Case QD_INCREASEICONINDEX
            If hDC = 0 And .IconTilesetName <> "" Then
                Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, .IconTileIndex + .QuantityOwned, .DisplayX, .DisplayY, 0, 0, 640, 480, False
            End If
        Case QD_ICONINDEXREPEATRIGHT
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To -Int(-.MaxQuantity / .IconCountPerRepeat) - 1
                    If .QuantityOwned <= .IconCountPerRepeat * Idx Then
                        DrawIdx = 0
                    ElseIf .QuantityOwned < .IconCountPerRepeat * (Idx + 1) Then
                        DrawIdx = .IconTileIndex + (.QuantityOwned Mod .IconCountPerRepeat) - 1
                    Else
                        DrawIdx = .IconTileIndex + .IconCountPerRepeat - 1
                    End If
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, DrawIdx, _
                        .DisplayX + Prj.TileSetDef(.IconTilesetIndex).TileWidth * Idx, .DisplayY, 0, 0, 640, 480, False
                Next
            End If
        Case QD_ICONINDEXREPEATABOVE
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To -Int(-.MaxQuantity / .IconCountPerRepeat) - 1
                    If .QuantityOwned <= .IconCountPerRepeat * Idx Then
                        DrawIdx = 0
                    ElseIf .QuantityOwned < .IconCountPerRepeat * (Idx + 1) Then
                        DrawIdx = .IconTileIndex + (.QuantityOwned Mod .IconCountPerRepeat) - 1
                    Else
                        DrawIdx = .IconTileIndex + .IconCountPerRepeat - 1
                    End If
                
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, DrawIdx, _
                        .DisplayX, .DisplayY - Prj.TileSetDef(.IconTilesetIndex).TileHeight * Idx, 0, 0, 640, 480, False
                Next
            End If
        Case QD_ICONINDEXREPEATBELOW
            If hDC = 0 And .IconTilesetName <> "" Then
                For Idx = 0 To -Int(-.MaxQuantity / .IconCountPerRepeat) - 1
                    If .QuantityOwned <= .IconCountPerRepeat * Idx Then
                        DrawIdx = 0
                    ElseIf .QuantityOwned < .IconCountPerRepeat * (Idx + 1) Then
                        DrawIdx = .IconTileIndex + (.QuantityOwned Mod .IconCountPerRepeat) - 1
                    Else
                        DrawIdx = .IconTileIndex + .IconCountPerRepeat - 1
                    End If
                    
                    Disp.DrawTile Prj.TileSetDef(.IconTilesetIndex).LoadedInstance, DrawIdx, _
                        .DisplayX, .DisplayY + Prj.TileSetDef(.IconTilesetIndex).TileHeight * Idx, 0, 0, 640, 480, False
                Next
            End If
        End Select
    End With
    
    Exit Sub
    
DrawInvError:
    ErrDesc = Err.Description
    If Not (Disp Is Nothing) Then
        Disp.Close
    ElseIf Not (CurDisp Is Nothing) Then
        CurDisp.Close
    End If
    Set CurDisp = Nothing
    MsgBox ErrDesc, vbExclamation, "Draw Inventory Item"
End Sub

Public Sub DrawInventory()
Attribute DrawInventory.VB_Description = "Draw all components of the inventory onto the screen in appropriate places."
    Dim Idx As Integer
    Dim hDC As Long
    
    For Idx = 0 To InvCount - 1
        DrawInvItem Idx
    Next
    hDC = Disp.GetDC
    For Idx = 0 To InvCount - 1
        DrawInvItem Idx, hDC
    Next
    Disp.ReleaseDC
End Sub

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the persistable contents of the game player object to a string for the project file."
    Dim Idx As Integer
    Dim Rslt As String

    Rslt = GeneralSerialize("ScrollMarginX", "ScrollMarginY", "InvMargin", "StartMap", "InvCount", _
                            "KeyUp", "KeyLeft", "KeyRight", "KeyDown", "KeyBtn1", "KeyBtn2", "KeyBtn3", _
                            "KeyBtn4", "EnableJoystick", _
                            ScrollMarginX, ScrollMarginY, InvBarMargin, StartMapName, InvCount, KeyConfig(0), _
                            KeyConfig(1), KeyConfig(2), KeyConfig(3), KeyConfig(4), KeyConfig(5), KeyConfig(6), _
                            KeyConfig(7), bEnableJoystick)
    For Idx = 0 To InvCount - 1
        With Inventory(Idx)
            Rslt = Rslt & GeneralSerialize( _
                "InvItemDef", "Name", "IconTilesetIndex", "IconTileIndex", "QuantityOwned", "MaxQuantity", _
                "QuantityDisplayType", "IconCountPerRepeat", "BarColor", "BarBackgroundColor", _
                "BarOutlineColor", "BarLength", "BarThickness", "DisplayX", "DisplayY", _
                Idx, .Name, .IconTilesetIndex, .IconTileIndex, .QuantityOwned, .MaxQuantity, _
                .QuantityDisplayType, .IconCountPerRepeat, .BarColor, .BarBackgroundColor, _
                .BarOutlineColor, .BarLength, .BarThickness, .DisplayX, .DisplayY)
        End With
    Next
        
    Serialize = Rslt
End Function

Public Sub Deserialize(Data As String)
Attribute Deserialize.VB_Description = "Load the player-related information from a project file string."
    Dim Idx As Integer
    Dim InvIdx As Integer
    Dim IntIdx As Integer
    Dim IntCount As Integer

    GeneralDeserialize Data, "ScrollMarginX", "ScrollMarginY", "InvMargin", "StartMap", "InvCount", _
                       "KeyUp", "KeyLeft", "KeyRight", "KeyDown", "KeyBtn1", "KeyBtn2", "KeyBtn3", _
                       "KeyBtn4", "EnableJoystick", _
                       ScrollMarginX, ScrollMarginY, InvBarMargin, StartMapName, InvCount, KeyCfg(0), _
                       KeyCfg(1), KeyCfg(2), KeyCfg(3), KeyCfg(4), KeyCfg(5), KeyCfg(6), _
                       KeyCfg(7), bEnableJoystick

    If InvCount > 0 Then
        ReDim Inventory(InvCount - 1)
    Else
        Erase Inventory
    End If
    
    For Idx = 0 To InvCount - 1
        GeneralDeserialize Data, "InvItemDef", InvIdx
        With Inventory(InvIdx)
            GeneralDeserialize Data, "Name", "IconTilesetIndex", "IconTileIndex", "QuantityOwned", "MaxQuantity", _
                               "QuantityDisplayType", "IconCountPerRepeat", "BarColor", "BarBackgroundColor", _
                               "BarOutlineColor", "BarLength", "BarThickness", "DisplayX", "DisplayY", _
                               .Name, .IconTilesetIndex, .IconTileIndex, .QuantityOwned, .MaxQuantity, _
                               .QuantityDisplayType, .IconCountPerRepeat, .BarColor, .BarBackgroundColor, _
                               .BarOutlineColor, .BarLength, .BarThickness, .DisplayX, .DisplayY
            .IconTilesetName = Prj.TileSetDef(.IconTilesetIndex).Name
        End With
    Next

End Sub

Public Sub Play(ByVal ScreenDepth As Integer)
Attribute Play.VB_Description = "The main play loop initializes and plays the game after it's loaded."
    Dim ErrStr As String
    Dim Idx

    On Error GoTo PlayErr

    ReIndexTilesetRefs
    If CurDisp Is Nothing Then
        Set CurDisp = New BMDXDisplay
        On Error Resume Next
        CurDisp.ValidateLicense "bygLILqJJySSOonPmqAZGuZp"
        On Error GoTo PlayErr
        CurDisp.OpenEx , , ScreenDepth
    End If
    Set rMap = Prj.Maps(StartMapName)
    Set Disp = CurDisp
    Set rMap.Disp = Disp
    rMap.LoadTiles
    rMap.LoadTileAnims
    rMap.InitSprites
    If Prj.MediaMgr.ClipExists(rMap.BackgroundMusic) Then
        Prj.MediaMgr.Clip(rMap.BackgroundMusic).Play
    End If
    If Not InitPlayerSprite Then
        Set CurDisp = Nothing
        rMap.Disp.Close
        Set rMap.Disp = Nothing
        MsgBox "Player sprite not found", vbExclamation
        Exit Sub
    End If

    bQuit = False
    Disp.Cls
    Set Disp = CurDisp
    For Idx = 0 To rMap.SpecialCount - 1
        If rMap.Specials(Idx).Flags And InteractionFlags.INTFL_ACTONSTARTUP Then
            ActivateFunction rMap.Specials(Idx)
        End If
    Next
    RaiseEvent OnPlayInit
    Do
        DoEvents
        If Not GameHost Is Nothing Then
            GameHost.RunStartScript
            If GameHost.CheckForError Then Exit Do
        End If
        If bEnableJoystick Then
            PollJoystick JOYSTICKID1
            If JSValid Then JoyMove
        Else
            JSValid = False
        End If
        CtlActions = KeyActions Or JoyActions
        AutoScrollMap
        DrawMap
        If rMap.bActive Then
            DrawInventory
            RaiseEvent OnBeforeMoveSprites
            rMap.AdvanceSprites
            RaiseEvent OnAfterMoveSprites
            Touch.Move PlayerSprite.X, PlayerSprite.Y
            Touch.TouchTest PlayerSprite, InterestingTiles, FirstTouchTiles
            rMap.CheckSpriteCollisions
            Prj.MediaMgr.RenderActiveMedia
        End If
        Disp.Flip
    Loop Until bQuit
    Set Disp = Nothing
    Set Touch = Nothing
    
    rMap.Disp.Close
    If Not GameHost Is Nothing Then GameHost.CheckForError
    Set rMap.Disp = Nothing
    Set CurDisp = Nothing
    Set rMap = Nothing
    Prj.MediaMgr.ModalFadeOutAll
    Prj.MediaMgr.UnloadAll
    RaiseEvent OnPlayComplete
    If Not GameHost Is Nothing Then
        GameHost.RunStartScript
        GameHost.CheckForError
    End If
    Exit Sub

PlayErr:
    ErrStr = Err.Description
    bQuit = True
    If Not (CurDisp Is Nothing) Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    Set Disp = Nothing
    If Not (rMap Is Nothing) Then
        Set rMap.Disp = Nothing
        Set rMap = Nothing
    End If
    MsgBox "Error playing map: " & ErrStr, vbCritical
    If Not GameHost Is Nothing Then GameHost.CheckForError
End Sub

Public Sub DrawMap()
Attribute DrawMap.VB_Description = "Draw the map according to its current scrolled position in the game player."
    rMap.Draw MapScrollX, MapScrollY, True
End Sub

Public Sub JoyMove()
Attribute JoyMove.VB_Description = "Process the current state of the joystick and raises OnControllerMove events if appropriate."
    Dim PrevState As Integer
    Dim ButtonMask As Integer
    
    PrevState = JoyActions
    ButtonMask = eActionBits.ACTION_BUTTON1 Or eActionBits.ACTION_BUTTON2 Or eActionBits.ACTION_BUTTON3 Or eActionBits.ACTION_BUTTON4
    
    If JoyX < 10000 Then JoyActions = JoyActions Or eActionBits.ACTION_LEFT Else JoyActions = JoyActions And Not eActionBits.ACTION_LEFT
    If JoyX > 55000 Then JoyActions = JoyActions Or eActionBits.ACTION_RIGHT Else JoyActions = JoyActions And Not eActionBits.ACTION_RIGHT
    If JoyY < 10000 Then JoyActions = JoyActions Or eActionBits.ACTION_UP Else JoyActions = JoyActions And Not eActionBits.ACTION_UP
    If JoyY > 55000 Then JoyActions = JoyActions Or eActionBits.ACTION_DOWN Else JoyActions = JoyActions And Not eActionBits.ACTION_DOWN
    If JoyButton1 Then JoyActions = JoyActions Or eActionBits.ACTION_BUTTON1 Else JoyActions = JoyActions And Not eActionBits.ACTION_BUTTON1
    If JoyButton2 Then JoyActions = JoyActions Or eActionBits.ACTION_BUTTON2 Else JoyActions = JoyActions And Not eActionBits.ACTION_BUTTON2
    If JoyButton3 Then JoyActions = JoyActions Or eActionBits.ACTION_BUTTON3 Else JoyActions = JoyActions And Not eActionBits.ACTION_BUTTON3
    If JoyButton4 Then JoyActions = JoyActions Or eActionBits.ACTION_BUTTON4 Else JoyActions = JoyActions And Not eActionBits.ACTION_BUTTON4
    If (JoyActions And ButtonMask) > (PrevState And ButtonMask) Then
        bPause = False
        Prj.MediaMgr.CancelModalMedia
    End If
    If JoyActions <> PrevState Then
        RaiseEvent OnControllerMove(PrevState, JoyActions)
    End If
End Sub

Public Sub AutoScrollMap()
Attribute AutoScrollMap.VB_Description = "Scroll the map according to the player's position."
    If PlayerSprite.X < MapScrollX + ScrollMarginX Then
        MapScrollX = PlayerSprite.X - ScrollMarginX
    ElseIf PlayerSprite.X + PlayerSprite.Width > MapScrollX + rMap.ViewWidth - ScrollMarginX Then
        MapScrollX = PlayerSprite.X + PlayerSprite.Width - rMap.ViewWidth + ScrollMarginX
    End If
    
    If PlayerSprite.Y < MapScrollY + ScrollMarginY Then
        MapScrollY = PlayerSprite.Y - ScrollMarginY
    ElseIf PlayerSprite.Y + PlayerSprite.Height > MapScrollY + rMap.ViewHeight - ScrollMarginY Then
        MapScrollY = PlayerSprite.Y + PlayerSprite.Height - rMap.ViewHeight + ScrollMarginY
    End If
    If MapScrollX > rMap.MapWidth - rMap.ViewWidth Then
        MapScrollX = rMap.MapWidth - rMap.ViewWidth
    End If
    If MapScrollY > rMap.MapHeight - rMap.ViewHeight Then
        MapScrollY = rMap.MapHeight - rMap.ViewHeight
    End If
    If MapScrollX < 0 Then MapScrollX = 0
    If MapScrollY < 0 Then MapScrollY = 0
End Sub

Private Sub Class_Initialize()
    KeyConfig(0) = vbKeyUp
    KeyConfig(1) = vbKeyLeft
    KeyConfig(2) = vbKeyRight
    KeyConfig(3) = vbKeyDown
    KeyConfig(4) = vbKeyControl
    KeyConfig(5) = vbKeySpace
    KeyConfig(6) = vbKeyShift
    KeyConfig(7) = vbKeyMenu
End Sub

Private Sub Disp_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim PrevState As Integer
    Dim ButtonMask As Integer

    PrevState = KeyActions
    ButtonMask = eActionBits.ACTION_BUTTON1 Or eActionBits.ACTION_BUTTON2 Or eActionBits.ACTION_BUTTON3 Or eActionBits.ACTION_BUTTON4
    
    If Shift = 0 And (KeyCode = vbKeyQ Or KeyCode = vbKeyEscape) Then bQuit = True
    If KeyCode = KeyConfig(0) Then KeyActions = KeyActions Or eActionBits.ACTION_UP
    If KeyCode = KeyConfig(1) Then KeyActions = KeyActions Or eActionBits.ACTION_LEFT
    If KeyCode = KeyConfig(2) Then KeyActions = KeyActions Or eActionBits.ACTION_RIGHT
    If KeyCode = KeyConfig(3) Then KeyActions = KeyActions Or eActionBits.ACTION_DOWN
    If KeyCode = KeyConfig(4) Then KeyActions = KeyActions Or eActionBits.ACTION_BUTTON1
    If KeyCode = KeyConfig(5) Then KeyActions = KeyActions Or eActionBits.ACTION_BUTTON2
    If KeyCode = KeyConfig(6) Then KeyActions = KeyActions Or eActionBits.ACTION_BUTTON3
    If KeyCode = KeyConfig(7) Then KeyActions = KeyActions Or eActionBits.ACTION_BUTTON4
    If (KeyActions And ButtonMask) > (PrevState And ButtonMask) Then
        bPause = False
        Prj.MediaMgr.CancelModalMedia
    End If
    If KeyActions <> PrevState Then
        RaiseEvent OnControllerMove(PrevState, KeyActions)
    End If
End Sub

Public Sub PollJoystick(ByVal StickNum As Integer)
Attribute PollJoystick.VB_Description = "Read the current state of the joystick."
    JSInfo.dwSize = Len(JSInfo)
    JSInfo.dwFlags = JOY_RETURNALL
    If joyGetPosEx(StickNum, JSInfo) = JOYERR_NOERROR Then
        JSValid = True
    Else
        JSValid = False
    End If
End Sub

Public Property Get JoyX() As Long
Attribute JoyX.VB_Description = "Return the X coordinate of the joystick (0-65535) when PollJoystick was last called."
    JoyX = JSInfo.dwXpos
End Property

Public Property Get JoyY() As Long
Attribute JoyY.VB_Description = "Return the X coordinate of the joystick (0-65535) when PollJoystick was last called."
    JoyY = JSInfo.dwYpos
End Property

Public Property Get JoyButton1() As Boolean
Attribute JoyButton1.VB_Description = "Return True if button 1 on the joystick was pressed when PollJoystick was last called."
    JoyButton1 = (JSInfo.dwButtons And JOY_BUTTON1) <> 0
End Property

Public Property Get JoyButton2() As Boolean
Attribute JoyButton2.VB_Description = "Return True if button 2 on the joystick was pressed when PollJoystick was last called."
    JoyButton2 = (JSInfo.dwButtons And JOY_BUTTON2) <> 0
End Property

Public Property Get JoyButton3() As Boolean
Attribute JoyButton3.VB_Description = "Return True if button 3 on the joystick was pressed when PollJoystick was last called."
    JoyButton3 = (JSInfo.dwButtons And JOY_BUTTON3) <> 0
End Property

Public Property Get JoyButton4() As Boolean
Attribute JoyButton4.VB_Description = "Return True if button 4 on the joystick was pressed when PollJoystick was last called."
    JoyButton4 = (JSInfo.dwButtons And JOY_BUTTON4) <> 0
End Property

Public Function InitPlayerSprite() As Boolean
Attribute InitPlayerSprite.VB_Description = "Initialize FirstTouchTiles, InterestingTiles and PlayerSprite properties."
    Dim LyrIdx As Integer
    Dim SprIdx As Integer
    Dim Idx As Integer
    
    Set Touch = New MapInteract
    For LyrIdx = 0 To rMap.LayerCount - 1
        For SprIdx = 0 To rMap.MapLayer(LyrIdx).SpriteCount - 1
            If rMap.MapLayer(LyrIdx).Sprite(SprIdx).rDef.Name = rMap.PlayerSpriteName Then
                Set PlayerSprite = rMap.MapLayer(LyrIdx).Sprite(SprIdx)
                InitPlayerSprite = True
                LyrIdx = rMap.LayerCount
                Exit For
            End If
        Next
    Next
    Set InterestingTiles = New TileGroup
    Set FirstTouchTiles = New TileGroup
    
    If Not (PlayerSprite Is Nothing) Then
        With PlayerSprite.rDef.rLayer.pMap
            For Idx = 0 To .InteractCount - 1
                InterestingTiles.MergeFrom .Interactions(Idx).TouchCategory.Group
                If .Interactions(Idx).Flags And InteractionFlags.INTFL_INITIALTOUCH Then
                    FirstTouchTiles.MergeFrom .Interactions(Idx).TouchCategory.Group
                End If
            Next
        End With
    End If
    rMap.bDisablePlayerEdit = False
End Function

Private Sub Disp_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim PrevState As Integer
    
    PrevState = KeyActions
    
    If KeyCode = KeyConfig(0) Then KeyActions = KeyActions And Not eActionBits.ACTION_UP
    If KeyCode = KeyConfig(1) Then KeyActions = KeyActions And Not eActionBits.ACTION_LEFT
    If KeyCode = KeyConfig(2) Then KeyActions = KeyActions And Not eActionBits.ACTION_RIGHT
    If KeyCode = KeyConfig(3) Then KeyActions = KeyActions And Not eActionBits.ACTION_DOWN
    If KeyCode = KeyConfig(4) Then KeyActions = KeyActions And Not eActionBits.ACTION_BUTTON1
    If KeyCode = KeyConfig(5) Then KeyActions = KeyActions And Not eActionBits.ACTION_BUTTON2
    If KeyCode = KeyConfig(6) Then KeyActions = KeyActions And Not eActionBits.ACTION_BUTTON3
    If KeyCode = KeyConfig(7) Then KeyActions = KeyActions And Not eActionBits.ACTION_BUTTON4
    If KeyActions <> PrevState Then
        RaiseEvent OnControllerMove(PrevState, KeyActions)
    End If
End Sub

Public Sub PlayFunctionMedia(SpecialObj As SpecialFunction)
Attribute PlayFunctionMedia.VB_Description = "Play a media clip associated with a function."
    If Len(SpecialObj.MediaName) > 0 Then
        If SpecialObj.Flags And InteractionFlags.INTFL_CHANGEBGMEDIA Then
            With PlayerSprite.rDef.rLayer.pMap
                Prj.MediaMgr.Clip(.BackgroundMusic).Play SpecialObj.MediaName
            End With
        Else
            If Prj.MediaMgr.ClipExists(SpecialObj.MediaName) Then
                If SpecialObj.Flags And InteractionFlags.INTFL_STOPMEDIA Then
                    Prj.MediaMgr.Clip(SpecialObj.MediaName).FadeOut
                Else
                    Prj.MediaMgr.Clip(SpecialObj.MediaName).Play
                End If
            End If
        End If
    End If
End Sub

Public Sub ActivateFunction(SpecialObj As SpecialFunction)
Attribute ActivateFunction.VB_Description = "Perform the actions dictated by a SpecialFunction."
    Dim NewSpr As Sprite
    Dim TmpCtl As Integer
    Dim Idx As Integer
    Dim Idx2 As Integer
    Dim hDC As Long
    Dim rcTextRect As RECT
    Dim hprPen As Long
    Dim hprBrush As Long
    Dim hprFont As Long
    Dim OtherFunc As SpecialFunction
    Dim MsgPic As StdPicture
    
    On Error GoTo SpecialErr
        
    If SpecialObj.FuncType = SPECIAL_MESSAGE Then
        If UCase$(Left$(SpecialObj.Value, 4)) <> "#PIC" Then
            hDC = Disp.GetDC(1)
            hprPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
            hprBrush = SelectObject(hDC, GetStockObject(BLACK_BRUSH))
            hprFont = SelectObject(hDC, GetStockObject(DEFAULT_GUI_FONT))
            With rcTextRect
                .Left = 160
                .Top = 0
                .Right = 480
                .Bottom = 0
                DrawText hDC, SpecialObj.Value, Len(SpecialObj.Value), rcTextRect, DT_CALCRECT Or DT_WORDBREAK
                .Top = 240 - .Bottom \ 2
                .Bottom = .Bottom + .Top
            End With
            Rectangle hDC, rcTextRect.Left - 2, rcTextRect.Top - 1, rcTextRect.Right + 2, rcTextRect.Bottom + 2
            SetTextColor hDC, vbWhite
            SetBkColor hDC, vbBlack
            DrawText hDC, SpecialObj.Value, Len(SpecialObj.Value), rcTextRect, DT_WORDBREAK
            SelectObject hDC, hprPen
            SelectObject hDC, hprBrush
            SelectObject hDC, hprFont
            hDC = 0
            Disp.ReleaseDC 1
        Else
            ChDir PathFromFile(Prj.ProjectPath)
            ChDrive Left$(PathFromFile(Prj.ProjectPath), 1)
            Set MsgPic = LoadPicture(Trim$(Mid$(SpecialObj.Value, 5)))
            Disp.PaintPicture MsgPic, 320 - Disp.HimetricToPixelX(MsgPic.Width) \ 2, 240 - Disp.HimetricToPixelY(MsgPic.Height) \ 2
            Disp.Flip
            Set MsgPic = Nothing
        End If
        
        PlayFunctionMedia SpecialObj
        bPause = True
        Do While bPause
            Prj.MediaMgr.RenderActiveMedia
            If bEnableJoystick Then
                PollJoystick JOYSTICKID1
                If JSValid Then JoyMove
            End If
            DoEvents
            Sleep 100
        Loop
        Disp.Flip
        Disp.Cls
        DrawMap
        DrawInventory
    ElseIf SpecialObj.FuncType = SPECIAL_SWITCHSPRITE Then
        PlayFunctionMedia SpecialObj
        With PlayerSprite.rDef.rLayer
            If SpecialObj.Flags And InteractionFlags.INTFL_NEWINSTANCE Then
                Set NewSpr = .pMap.SpriteDefs(SpecialObj.SpriteName).MakeInstance
                .AddSprite NewSpr
            Else
                For Idx = 0 To .SpriteCount - 1
                    If .Sprite(Idx).rDef.Name = SpecialObj.SpriteName Then
                        Set NewSpr = .Sprite(Idx)
                        Exit For
                    End If
                Next
                If Idx >= .SpriteCount Then
                    Set NewSpr = .pMap.SpriteDefs(SpecialObj.SpriteName).MakeInstance
                    .AddSprite NewSpr
                End If
            End If
            If SpecialObj.Flags And InteractionFlags.INTFL_OLDLOCATION Then
                NewSpr.X = PlayerSprite.X
                NewSpr.Y = PlayerSprite.Y
            End If
            If SpecialObj.Flags And InteractionFlags.INTFL_SWAPCONTROL Then
                TmpCtl = NewSpr.rDef.Template.ControlType
                NewSpr.rDef.Template.ControlType = PlayerSprite.rDef.Template.ControlType
                PlayerSprite.rDef.Template.ControlType = TmpCtl
            End If
            If SpecialObj.Flags And InteractionFlags.INTFL_DELETEOLD Then
                For Idx = 0 To .SpriteCount - 1
                    If .Sprite(Idx) Is PlayerSprite Then
                        .RemoveSprite Idx
                        Exit For
                    End If
                Next
            End If
            Set PlayerSprite = NewSpr
        End With
    ElseIf SpecialObj.FuncType = SPECIAL_SWITCHMAP Then
        PlayFunctionMedia SpecialObj
        Set rMap.Disp = Nothing
        TmpCtl = rMap.bDisablePlayerEdit
        Set rMap = Prj.Maps(SpecialObj.Value)
        Set rMap.Disp = Disp
        rMap.LoadTiles
        rMap.LoadTileAnims
        rMap.InitSprites
        InitPlayerSprite
        If Len(SpecialObj.SpriteName) > 0 Then
            For Idx = 0 To rMap.LayerCount - 1
                With rMap.MapLayer(Idx)
                    For Idx2 = 0 To .SpriteCount - 1
                        If .Sprite(Idx2).rDef.Name = SpecialObj.Name Then
                            Set PlayerSprite = .Sprite(Idx2)
                            Idx = rMap.LayerCount
                            Exit For
                        End If
                    Next
                End With
            Next
        End If
        rMap.bDisablePlayerEdit = TmpCtl

        If SpecialObj.Flags And InteractionFlags.INTFL_OVERRIDEPOSITION Then
            PlayerSprite.X = SpecialObj.DestX
            PlayerSprite.Y = SpecialObj.DestY
        End If
        
        Prj.MediaMgr.ModalFadeOutAll
        If Prj.MediaMgr.ClipExists(rMap.BackgroundMusic) Then
            Prj.MediaMgr.Clip(rMap.BackgroundMusic).Play
        End If
    ElseIf SpecialObj.FuncType = SPECIAL_TELEPORT Then
        PlayFunctionMedia SpecialObj
        If SpecialObj.Flags And InteractionFlags.INTFL_RELATIVETELEPORT Then
            PlayerSprite.X = PlayerSprite.X + SpecialObj.DestX
            PlayerSprite.Y = PlayerSprite.Y + SpecialObj.DestY
        Else
            PlayerSprite.X = SpecialObj.DestX
            PlayerSprite.Y = SpecialObj.DestY
        End If
    ElseIf SpecialObj.FuncType = SPECIAL_ALTERMAP Then
        PlayFunctionMedia SpecialObj
        With PlayerSprite.rDef.rLayer.pMap
            If Not .bDisablePlayerEdit Then
                If .SpecialExists(SpecialObj.Value) Then
                    Set OtherFunc = .Specials(SpecialObj.Value)
                    With .MapLayer(SpecialObj.LayerIndex)
                        For Idx = 0 To OtherFunc.TileBottom - OtherFunc.TileTop
                            If OtherFunc.DestY + Idx >= 0 And OtherFunc.DestY + Idx < .Rows Then
                                For Idx2 = 0 To OtherFunc.TileRight - OtherFunc.TileLeft
                                    If OtherFunc.DestX + Idx2 >= 0 And OtherFunc.DestX + Idx2 < .Columns Then
                                        .Data.TileValue(SpecialObj.DestX + Idx2, SpecialObj.DestY + Idx) = .Data.TileValue(OtherFunc.TileLeft + Idx2, OtherFunc.TileTop + Idx)
                                    End If
                                Next
                            End If
                        Next
                    End With
                End If
            End If
        End With
    ElseIf SpecialObj.FuncType = SPECIAL_CREATESPRITE Then
        PlayFunctionMedia SpecialObj
        'With PlayerSprite.rDef.rLayer
        With PlayerSprite.rDef.rLayer.pMap.SpriteDefs(SpecialObj.SpriteName)
            Idx2 = 0
            If Val(SpecialObj.Value) > 0 Then
                For Idx = 0 To .rLayer.SpriteCount - 1
                    If .rLayer.Sprite(Idx).rDef.Name = .Name Then
                        Idx2 = Idx2 + 1
                        If Idx2 >= Val(SpecialObj.Value) Then Exit For
                    End If
                Next
            End If
            If (Idx2 < Val(SpecialObj.Value)) Or (Val(SpecialObj.Value) = 0) Then
                Set NewSpr = .MakeInstance
                ' I see no reason to force the new sprite onto the player's layer
                '.AddSprite NewSpr
                .rLayer.AddSprite NewSpr
                With NewSpr
                    If SpecialObj.Flags And InteractionFlags.INTFL_OVERRIDEPOSITION Then
                        If SpecialObj.Flags And InteractionFlags.INTFL_RELATIVETOPLAYER Then
                            .X = PlayerSprite.X + SpecialObj.DestX
                            .Y = PlayerSprite.Y + SpecialObj.DestY
                        Else
                            .X = SpecialObj.DestX
                            .Y = SpecialObj.DestY
                        End If
                    End If
                End With
            End If
        End With
    ElseIf SpecialObj.FuncType = SPECIAL_DELETESPRITE Then
        PlayFunctionMedia SpecialObj
        With PlayerSprite.rDef.rLayer.pMap.SpriteDefs(SpecialObj.SpriteName)
            Idx2 = 0
            For Idx = .rLayer.SpriteCount - 1 To 0 Step -1
                If .rLayer.Sprite(Idx).rDef.Name = .Name Then
                    Idx2 = Idx2 + 1
                    If Idx2 > Val(SpecialObj.Value) Then
                        .rLayer.RemoveSprite Idx
                        If (SpecialObj.Flags And InteractionFlags.INTFL_DELETEMANY) = 0 Then Exit For
                    End If
                End If
            Next
        End With
    Else
        PlayFunctionMedia SpecialObj
    End If

    If SpecialObj.Flags And InteractionFlags.INTFL_RAISEEVENT Then
        RaiseEvent OnSpecialFunction(SpecialObj)
    End If
    If SpecialObj.Flags And InteractionFlags.INTFL_REMOVEALWAYS Then
        If Not PlayerSprite.rDef.rLayer.pMap.bDisablePlayerEdit Then
            PlayerSprite.rDef.rLayer.pMap.RemoveSpecial SpecialObj.Name
        End If
    End If
    
    Exit Sub

SpecialErr:
    Dim ErrMsg As String
    ErrMsg = Err.Description
    If Not (Disp Is Nothing) Then
        Disp.Close
        Set Disp = Nothing
        Set CurDisp = Nothing
    ElseIf Not (CurDisp Is Nothing) Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    MsgBox ErrMsg, vbExclamation
End Sub

Private Sub Touch_OnTouchSpecial(SpecialObj As SpecialFunction)
    Dim ActionFlags As Integer
    
    If CtlActions And eActionBits.ACTION_BUTTON1 Then ActionFlags = InteractionFlags.INTFL_ACTONBUTTON
    If CtlActions And eActionBits.ACTION_DOWN Then ActionFlags = ActionFlags Or InteractionFlags.INTFL_ACTONDOWN
    If CtlActions And eActionBits.ACTION_UP Then ActionFlags = ActionFlags Or InteractionFlags.INTFL_ACTONUP

    If ((SpecialObj.Flags And ActionFlags) = 0) Then
        If SpecialObj.Flags And (InteractionFlags.INTFL_ACTONDOWN Or _
            InteractionFlags.INTFL_ACTONBUTTON Or InteractionFlags.INTFL_ACTONUP) Then Exit Sub
    End If
    
    If SpecialObj.InvUseCount > 0 Then
        If InvQuantityOwned(SpecialObj.InvItem) < SpecialObj.InvUseCount Then
            Exit Sub
        Else
            If SpecialObj.Flags And InteractionFlags.INTFL_FUNCREMOVEINV Then
                If Not rMap.bDisablePlayerEdit Then
                    InvQuantityOwned(SpecialObj.InvItem) = InvQuantityOwned(SpecialObj.InvItem) - SpecialObj.InvUseCount
                End If
            End If
        End If
    End If
    
    ActivateFunction SpecialObj
    
End Sub

Private Sub Touch_OnTouchTile(ByVal TileX As Long, ByVal TileY As Long, ByVal TileVal As Integer, ByVal bNewTouch As Boolean)
    Dim Idx As Integer
    Dim bReacted As Boolean
    
    On Error GoTo TouchErr

    If rMap.bDisablePlayerEdit Then Exit Sub

    For Idx = 0 To PlayerSprite.rDef.rLayer.pMap.InteractCount - 1
        With PlayerSprite.rDef.rLayer.pMap.Interactions(Idx)
            If .TouchCategory.Group.IsMember(TileVal) Then
                bReacted = False
                If .Reaction = REACT_ADDTOINVENTORY Then
                    With Inventory(.InvItem)
                        If .QuantityOwned < .MaxQuantity Then
                            .QuantityOwned = .QuantityOwned + 1
                            bReacted = True
                        End If
                    End With
                ElseIf .Reaction = REACT_SUBTRACTFROMINVENTORY Then
                    With Inventory(.InvItem)
                        If .QuantityOwned > 0 Then
                            .QuantityOwned = .QuantityOwned - 1
                            bReacted = True
                        End If
                    End With
                ElseIf .Reaction = REACT_USEINVENTORY Then
                    If Inventory(.InvItem).QuantityOwned > 0 Then
                        bReacted = True
                    End If
                Else
                    bReacted = True
                End If
                If (.Flags And InteractionFlags.INTFL_REMOVEIFACT And bReacted) Or (.Flags And InteractionFlags.INTFL_REMOVEALWAYS) Then
                    PlayerSprite.rDef.rLayer.Data.TileValue(TileX, TileY) = .ReplaceTile
                    bReacted = True
                End If
                If bReacted And Len(.Media) > 0 Then Prj.MediaMgr.Clip(.Media).Play
                If .Flags And InteractionFlags.INTFL_RAISEEVENT Then
                    RaiseEvent OnTileInteraction(TileX, TileY, TileVal, bNewTouch, Idx)
                End If
            End If
        End With
    Next

    Exit Sub
    
TouchErr:
    Dim ErrMsg As String
    ErrMsg = Err.Description
    If Not (Disp Is Nothing) Then
        Disp.Close
        Set Disp = Nothing
        Set CurDisp = Nothing
    ElseIf Not (CurDisp Is Nothing) Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    MsgBox ErrMsg, vbExclamation
End Sub

Public Sub TriggerCollide(ByVal Name As String, ByVal ClsASprIdx As Integer, ByVal ClsBSprIdx As Integer, ByVal CollDefIdx As Integer)
Attribute TriggerCollide.VB_Description = "Fires OnSpritesCollide"
    RaiseEvent OnSpritesCollide(Name, ClsASprIdx, ClsBSprIdx, CollDefIdx)
End Sub

Public Sub TriggerNewSprite(NewSprite As Sprite)
Attribute TriggerNewSprite.VB_Description = "Fires OnNewSprite."
    RaiseEvent OnNewSprite(NewSprite)
End Sub
