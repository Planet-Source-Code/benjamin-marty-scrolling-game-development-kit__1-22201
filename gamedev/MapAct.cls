VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapInteract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Code used to detect interactions between a sprite and the map.  No persistent data, one instance owned by project's GamePlayer (the only sprite that can interact by default)."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MapAct.cls - Sprite-Map Interaction Handling Class
'
'======================================================================

Option Explicit

Public LastX As Long
Attribute LastX.VB_VarProcData = ";Data"
Attribute LastX.VB_VarDescription = "Previous X position of the sprite being tested."
Public LastY As Long
Attribute LastY.VB_VarProcData = ";Data"
Attribute LastY.VB_VarDescription = "Previous Y position of the sprite being tested."
Public CurX As Long
Attribute CurX.VB_VarProcData = ";Data"
Attribute CurX.VB_VarDescription = "Current X position of the sprite being tested."
Public CurY As Long
Attribute CurY.VB_VarProcData = ";Data"
Attribute CurY.VB_VarDescription = "Current Y position of the sprite being tested."

Event OnTouchTile(ByVal TileX As Long, ByVal TileY As Long, ByVal TileVal As Integer, ByVal bNewTouch As Boolean)
Attribute OnTouchTile.VB_Description = "Fired when the sprite touches a tile according to specified conditions."
Event OnTouchSpecial(SpecialObj As SpecialFunction)
Attribute OnTouchSpecial.VB_Description = "Fired when the sprite touches a special function according to specified conditions."

Private Sub Class_Initialize()
    LastX = -1
    LastY = -1
End Sub

Public Sub TouchTest(SpriteObj As Sprite, Optional InterestingTiles As TileGroup = Nothing, Optional FirstTouchOnly As TileGroup = Nothing)
Attribute TouchTest.VB_Description = "Perform the tests and raise events according to specified parameters."
    Dim XI As Long
    Dim YI As Long
    Dim SprWid As Long
    Dim SprHgt As Long
    Dim TilWid As Long
    Dim TilHgt As Long
    Dim MinX As Long
    Dim MaxX As Long
    Dim MinY As Long
    Dim MaxY As Long
    Dim LastMinX As Long
    Dim LastMinY As Long
    Dim LastMaxX As Long
    Dim LastMaxY As Long
    Dim bNewTouch As Boolean
    Dim bInteresting As Boolean
    Dim TileVal As Integer
    Dim LayerObj As Layer
    
    Set LayerObj = SpriteObj.rDef.rLayer
    SprWid = SpriteObj.Width
    SprHgt = SpriteObj.Height
    TilWid = LayerObj.TSDef.TileWidth
    TilHgt = LayerObj.TSDef.TileHeight
    
    LastMinX = (LastX - 1) \ TilWid
    If LastMinX < 0 Then LastMinX = 0
    LastMinY = (LastY - 1) \ TilHgt
    If LastMinY < 0 Then LastMinY = 0
    LastMaxX = (LastX + SprWid) \ TilWid
    If LastMaxX >= LayerObj.Columns Then LastMaxX = LayerObj.Columns - 1
    LastMaxY = (LastY + SprHgt) \ TilHgt
    If LastMaxY >= LayerObj.Rows Then LastMaxY = LayerObj.Rows - 1
    
    MinX = (CurX - 1) \ TilWid
    If MinX < 0 Then MinX = 0
    MaxX = (CurX + SprWid) \ TilWid
    If MaxX >= LayerObj.Columns Then MaxX = LayerObj.Columns - 1
    For XI = MinX To MaxX
        MinY = (CurY - 1) \ TilHgt
        If MinY < 0 Then MinY = 0
        MaxY = (CurY + SprHgt) \ TilHgt
        If MaxY >= LayerObj.Rows Then MaxY = LayerObj.Rows - 1
        If (MinY + 1) * TilHgt <= CurY Then If ((XI + 1) * TilWid <= CurX Or XI * TilWid >= CurX + SprWid) Then MinY = MinY + 1
        If MaxY * TilHgt >= CurY + SprHgt Then If ((XI + 1) * TilWid <= CurX Or XI * TilWid >= CurX + SprWid) Then MaxY = MaxY - 1
        For YI = MinY To MaxY
            TileVal = LayerObj.Data.TileValue(XI, YI)
            If InterestingTiles Is Nothing Then
                bInteresting = True
            ElseIf InterestingTiles.IsMember(TileVal) Then
                bInteresting = True
            Else
                bInteresting = False
            End If
            If bInteresting Then
                bNewTouch = False
                If XI < LastMinX Or YI < LastMinY Or XI > LastMaxX Or YI > LastMaxY Then
                    bNewTouch = True
                ElseIf XI = LastMinX Or XI = LastMaxX Or YI = LastMinY Or YI = LastMaxY Then
                    ' On the border
                    If XI = LastMinX And YI = LastMinY Then
                        If (LastMinX + 1) * TilWid <= LastX And (LastMinY + 1) * TilHgt <= LastY Then bNewTouch = True
                    ElseIf XI = LastMaxX And YI = LastMinY Then
                        If LastMaxX * TilWid >= LastX + SprWid And (LastMinY + 1) * TilHgt <= LastY Then bNewTouch = True
                    ElseIf XI = LastMinX And YI = LastMaxY Then
                        If (LastMinX + 1) * TilWid <= LastX And LastMaxY * TilHgt >= LastY + SprHgt Then bNewTouch = True
                    ElseIf XI = LastMaxX And YI = LastMaxY Then
                        If LastMaxX * TilWid >= LastX + SprWid And LastMaxY * TilHgt >= LastY + SprHgt Then bNewTouch = True
                    End If ' Else old touch
                End If ' Else old touch
                If FirstTouchOnly Is Nothing Then
                    RaiseEvent OnTouchTile(XI, YI, TileVal, bNewTouch)
                ElseIf (Not FirstTouchOnly.IsMember(TileVal)) Or bNewTouch Then
                    RaiseEvent OnTouchTile(XI, YI, TileVal, bNewTouch)
                Else
                    bInteresting = False
                End If
            End If
        Next
    Next
    
    For XI = 0 To LayerObj.pMap.SpecialCount - 1
        If XI > LayerObj.pMap.SpecialCount - 1 Then Exit For
        With LayerObj.pMap.Specials(XI)
            If LayerObj.pMap.MapLayer(.LayerIndex) Is LayerObj Then
                If CurX <= (.TileRight + 1) * TilWid And CurY <= (.TileBottom + 1) * TilHgt And _
                    CurX + SprWid >= .TileLeft * TilWid And CurY + SprHgt >= .TileTop * TilHgt Then
                    If (CurX + SprWid > .TileLeft * TilWid) And (CurX < (.TileRight + 1) * TilWid) Then
                        bInteresting = True
                    ElseIf (CurY + SprHgt > .TileTop * TilHgt) And (CurY < (.TileBottom + 1) * TilHgt) Then
                        bInteresting = True
                    Else
                        bInteresting = False
                    End If
                    If bInteresting Then
                        If .Flags And InteractionFlags.INTFL_INITIALTOUCH Then
                            If LastX <= (.TileRight + 1) * TilWid And LastY <= (.TileBottom + 1) * TilHgt And _
                                LastX + SprWid >= .TileLeft * TilWid And LastY + SprHgt >= .TileTop * TilHgt Then
                                If (LastX + SprWid > .TileLeft * TilWid) And (LastX < (.TileRight + 1) * TilWid) Then
                                    bInteresting = False
                                ElseIf (LastY + SprHgt > .TileTop * TilHgt) And (LastY < (.TileBottom + 1) * TilHgt) Then
                                    bInteresting = False
                                Else
                                    bInteresting = True
                                End If
                            Else
                                bInteresting = True
                            End If
                        Else
                            bInteresting = True
                        End If
                        If bInteresting Then
                            RaiseEvent OnTouchSpecial(LayerObj.pMap.Specials(XI))
                        End If
                    End If
                End If
            End If
        End With
    Next
    
End Sub

Public Sub Move(ByVal NewX As Single, ByVal NewY As Single)
Attribute Move.VB_Description = "Notify this object that the sprite has moved. (Updates current and previous X,Y.)"
    LastX = CLng(CurX)
    LastY = CLng(CurY)
    CurX = CLng(NewX)
    CurY = CLng(NewY)
End Sub
