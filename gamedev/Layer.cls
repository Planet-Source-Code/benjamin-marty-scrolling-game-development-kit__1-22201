VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Layer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "One scrolling layer of a map.  Indexed and owned by map."
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
' File: Layer.cls - Map Layer Data Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this layer as indexed by the Map."
Public TSDef As TileSetDef
Attribute TSDef.VB_VarProcData = ";Relational"
Attribute TSDef.VB_VarDescription = "The TileSetDef object used to create and maintain the tileset for this layer's graphics."
Public Tileset As IBMDXTileSet
Attribute Tileset.VB_VarHelpID = -1
Attribute Tileset.VB_VarDescription = "The BMDXTileset object (display memory) containing this layer's tile graphics."
Private Width As Integer
Private Height As Integer
Public pMap As Map
Attribute pMap.VB_VarProcData = ";Relational"
Attribute pMap.VB_VarDescription = "Reference to the map in which this layer exists."
Public XScrollRate As Single
Attribute XScrollRate.VB_VarProcData = ";Behavior"
Attribute XScrollRate.VB_VarDescription = "Value indicating how fast this layer scrolls horizontally relative to the map (layer pixels per map pixel; may be non-integer)."
Public YScrollRate As Single
Attribute YScrollRate.VB_VarProcData = ";Behavior"
Attribute YScrollRate.VB_VarDescription = "Value indicating how fast this layer scrolls vertically relative to the map (layer pixels per map pixel; may be non-integer)."
Public Data As BMDXTileMap
Attribute Data.VB_VarUserMemId = 0
Attribute Data.VB_VarProcData = ";Data"
Attribute Data.VB_VarDescription = "Return the underlying BMDXMap object containing the tiles."
Public Transparent As Boolean
Attribute Transparent.VB_VarProcData = ";Appearance"
Attribute Transparent.VB_VarDescription = "Determines whether the background of the layer (background color of the tileset) is drawn."
Public TileAnimColl As New Collection
Attribute TileAnimColl.VB_VarDescription = "The collection of AnimDefs applicable to the layer, initialized with LoadTileAnims."
Private SpriteColl As Variant
Private ClsMbrCount(0 To 15) As Integer

Public Sub ReleaseRefs()
Attribute ReleaseRefs.VB_Description = "Release all references that the layer has to other objects to avoid circular reference problems."
    Set TSDef = Nothing
    Set Tileset = Nothing
    Set pMap = Nothing
    Set Data = Nothing
    Set TileAnimColl = Nothing
    SpriteColl = Empty
End Sub

Public Property Let ClsMbrCnt(ByVal Index As Integer, ByVal NewVal As Integer)
Attribute ClsMbrCnt.VB_Description = "How many sprites are in a collision class (for collision optimization). Set by Map.AdvanceSprites."
    ClsMbrCount(Index) = NewVal
End Property
Public Property Get ClsMbrCnt(ByVal Index As Integer) As Integer
    ClsMbrCnt = ClsMbrCount(Index)
End Property

Public Sub ClearSprites()
Attribute ClearSprites.VB_Description = "Remove all sprites from the layer."
    SpriteColl = Empty
End Sub

Sub Create(ByVal LayerName As String, ByVal PixWidth As Long, ByVal PixHeight As Long, TS As TileSetDef, ByVal XR As Single, ByVal YR As Single, ByVal Trans As Boolean, Mp As Map)
Attribute Create.VB_Description = "Initialize the layer object's essential data."
    Name = LayerName
    
    XScrollRate = XR
    YScrollRate = YR
    
    Width = -Int(-(PixWidth - 640) / TS.TileWidth * XScrollRate)
    If Width < 0 Then Width = 0
    Width = Width - Int(-640 / TS.TileWidth)
    
    If -Int(-PixWidth / TS.TileWidth * XScrollRate) > Width Then
        Width = -Int(-PixWidth / TS.TileWidth * XScrollRate)
    End If
    
    Height = -Int(-(PixHeight - 480) / TS.TileHeight * YScrollRate)
    If Height < 0 Then Height = 0
    Height = Height - Int(-480 / TS.TileHeight)
    
    If -Int(-PixHeight / TS.TileHeight * YScrollRate) > Height Then
        Height = -Int(-PixHeight / TS.TileHeight * YScrollRate)
    End If
    
    Transparent = Trans
    Set TSDef = TS

    Set Data = New BMDXTileMap
    Data.Create Width, Height
    
    Set pMap = Mp
End Sub

Property Get Columns() As Integer
Attribute Columns.VB_Description = "The width of the layer in tiles."
    Columns = Width
End Property

Property Get Rows() As Integer
Attribute Rows.VB_Description = "Return the height of the layer in tiles."
    Rows = Height
End Property

Sub LoadTiles(Disp As BMDXDisplay)
Attribute LoadTiles.VB_Description = "Load the graphics for this layer's tileset into display memory, or locate an existing copy."
    
    If Not TSDef.IsLoaded Then
        TSDef.Load
    End If
    If Not TSDef.IsLoaded Then
        Err.Raise vbObjectError, , "Failed to load " & TSDef.Name & " (" & TSDef.ImagePath & ")"
    End If
    
    If TSDef.LoadedInstance Is Nothing Then
        Set Tileset = Disp.CreateTileSet(TSDef.Image, TSDef.TileWidth, TSDef.TileHeight)
        Set TSDef.LoadedInstance = Tileset
    Else
        Set Tileset = TSDef.LoadedInstance
    End If
    
End Sub

Sub UnloadTiles()
Attribute UnloadTiles.VB_Description = "Releases the layer's reference to the BMDXTileset object."
    Set Tileset = Nothing
End Sub

Function IsLoaded() As Boolean
Attribute IsLoaded.VB_Description = "Returns True if the layer's tileset is loaded"
    If Tileset Is Nothing Then IsLoaded = False Else IsLoaded = True
End Function

Sub UpdateLayer(ByVal LayerName As String, ByVal PixWidth As Long, ByVal PixHeight As Long, TS As TileSetDef, ByVal XR As Single, ByVal YR As Single, ByVal Trans As Boolean)
Attribute UpdateLayer.VB_Description = "Update the layer to a new set of parameters, changing the size and cropping if necessary."
    Dim NewWidth As Integer
    Dim NewHeight As Integer
    Dim YIndex As Integer, XIndex As Integer
    Dim Store As Variant
    
    If IsLoaded() Then
        Err.Raise vbObjectError, , "Can't update layer while loaded"
    End If
    
    Name = LayerName
    XScrollRate = XR
    YScrollRate = YR
    Transparent = Trans
    Set TSDef = TS
    
    If PixWidth = 0 Then PixWidth = pMap.MapWidth
    If PixHeight = 0 Then PixHeight = pMap.MapHeight
    
    If PixWidth > 0 And PixHeight > 0 Then
        NewWidth = -Int(-(PixWidth - 640) / TS.TileWidth * XScrollRate)
        If NewWidth = 0 Then NewWidth = 0
        NewWidth = NewWidth - Int(-640 / TS.TileWidth)
        
        If -Int(-PixWidth / TS.TileWidth * XScrollRate) > NewWidth Then
            NewWidth = -Int(-PixWidth / TS.TileWidth * XScrollRate)
        End If
        
        NewHeight = -Int(-(PixHeight - 480) / TS.TileHeight * YScrollRate)
        If NewHeight < 0 Then NewHeight = 0
        NewHeight = NewHeight - Int(-480 / TS.TileHeight)
        
        If -Int(-PixHeight / TS.TileHeight * YScrollRate) > NewHeight Then
            NewHeight = -Int(-PixHeight / TS.TileHeight * YScrollRate)
        End If
        
        Store = Data.MapData
        Set Data = New BMDXTileMap
        Data.Create NewWidth, NewHeight
        
        For YIndex = 0 To NewHeight - 1
            For XIndex = 0 To NewWidth - 1
                If XIndex < Width And YIndex < Height Then
                    Data.TileValue(XIndex, YIndex) = Store(YIndex * Width + XIndex)
                End If
            Next
        Next
        
        Height = NewHeight
        Width = NewWidth
    End If
    
End Sub

Public Sub LoadTileAnims()
Attribute LoadTileAnims.VB_Description = "Initializes tile animation for the layer based on the project's AnimDefs."
    Dim I As Integer
    
    Do While TileAnimColl.Count
        TileAnimColl.Remove 1
    Loop
    
    For I = 0 To 255
        Data.TileMapping(I) = I
    Next
    
    For I = 0 To Prj.AnimDefCount - 1
        If Prj.AnimDefs(I).MapName = pMap.Name And Prj.AnimDefs(I).LayerName = Name Then
            TileAnimColl.Add Prj.AnimDefs(I)
            Prj.AnimDefs(I).Reset
        End If
    Next
    
End Sub

Public Sub AdvanceTileAnims()
Attribute AdvanceTileAnims.VB_Description = "Process all AnimDefs for this layer (animate the tiles)."
    Dim TA As AnimDef
    
    For Each TA In TileAnimColl
        TA.Advance
        Data.TileMapping(TA.BaseTile) = TA.CurTile
    Next
    
End Sub

Public Sub DrawLayerSprites(ByVal MapXOffset As Long, ByVal MapYOffset As Long)
Attribute DrawLayerSprites.VB_Description = "Draw all this layer's sprites to the map's current display according to the map's current scroll position."
   Dim I As Integer
   Dim Spr As Sprite
   
   For I = 0 To SpriteCount - 1
      Set Spr = SpriteColl(I)
      DrawSprite Spr, MapXOffset, MapYOffset
   Next
End Sub

Public Sub DrawSprite(Spr As Sprite, ByVal MapXOffset As Long, ByVal MapYOffset As Long)
Attribute DrawSprite.VB_Description = "Draw a single sprite object to this layer's map's display given the map's scroll position, accounting for the layer's scroll rate."
    Dim Tile As Integer, RX As Long, RY As Long
    Tile = Spr.CurTile
    RX = CLng(Spr.X - MapXOffset * XScrollRate + pMap.ViewLeft)
    RY = CLng(Spr.Y - MapYOffset * YScrollRate + pMap.ViewTop)
    If Tile >= 0 And Abs(RX) < 640 And Abs(RY) < 640 Then
        pMap.Disp.DrawTile Spr.CurTS, Spr.CurTile, RX, RY, pMap.ViewLeft, pMap.ViewTop, pMap.ViewWidth, pMap.ViewHeight, True
    End If
End Sub

Public Sub AddSprite(Spr As Sprite)
Attribute AddSprite.VB_Description = "Add a new Sprite object to this layer, to be drawn by DrawLayerSprites."
    Dim InsIdx As Integer
    If IsEmpty(SpriteColl) Then
        ReDim SpriteColl(0) As Sprite
    Else
        ReDim Preserve SpriteColl(UBound(SpriteColl) + 1) As Sprite
    End If
    InsIdx = UBound(SpriteColl)
    If Spr.rDef.Template.ControlType <> CONTROL_SIMPLE Then
        ' Keep all simple sprites at the end of the array
        Do
            If InsIdx = 0 Then Exit Do
            If SpriteColl(InsIdx - 1).rDef.Template.ControlType = CONTROL_SIMPLE Then
                Set SpriteColl(InsIdx) = SpriteColl(InsIdx - 1)
                InsIdx = InsIdx - 1
            Else
                Exit Do
            End If
        Loop
    End If
    Set SpriteColl(InsIdx) = Spr
    If Not Prj.GamePlayer Is Nothing Then
        If pMap.PlayerSpriteName = Spr.rDef.Name Then Set Prj.GamePlayer.PlayerSprite = Spr
    End If
    pMap.IsDirty = True
    If Not (Prj.GamePlayer Is Nothing) Then Prj.GamePlayer.TriggerNewSprite Spr
End Sub

Public Sub RemoveSprite(ByVal nIndex As Integer)
Attribute RemoveSprite.VB_Description = "Remove a sprite from the layer's collection of active sprites."
    Dim I As Integer
   
    For I = LBound(SpriteColl) To UBound(SpriteColl)
        If SpriteColl(I).pRideOnRef Is SpriteColl(nIndex) Then Set SpriteColl(I).pRideOnRef = Nothing
    Next
   
    For I = nIndex To UBound(SpriteColl) - 1
        Set SpriteColl(I) = SpriteColl(I + 1)
    Next
   
    Set SpriteColl(UBound(SpriteColl)) = Nothing
   
    If UBound(SpriteColl) > LBound(SpriteColl) Then
        ReDim Preserve SpriteColl(UBound(SpriteColl) - 1)
    Else
        SpriteColl = Empty
    End If
    pMap.IsDirty = True
End Sub

Public Function SpriteCount() As Integer
Attribute SpriteCount.VB_Description = "Return the number of active sprites in the layer."
   If IsEmpty(SpriteColl) Then
      SpriteCount = 0
   Else
      SpriteCount = UBound(SpriteColl) - LBound(SpriteColl) + 1
   End If
End Function

Public Property Get Sprite(ByVal nIndex As Integer) As Sprite
Attribute Sprite.VB_Description = "Return a sprite given its ordinal index in the layer."
   Set Sprite = SpriteColl(nIndex)
End Property

Public Sub CheckSpriteCollisions()
Attribute CheckSpriteCollisions.VB_Description = "Find, process and react to sprite collisions out of all this layer's sprites."
    Dim CollIdx As Integer
    Dim ClassA As Integer
    Dim ClassB As Integer
    Dim SprAIdx As Integer
    Dim SprBIdx As Integer
    Dim SprClsMsk As Integer
    Static Msk(0 To 15) As Integer
    Dim SCUBound As Integer
    Dim Spr As Sprite
    Dim bDoTest As Boolean

    If IsEmpty(SpriteColl) Then Exit Sub
    If Msk(0) = 0 Then
        For CollIdx = 0 To 14
            Msk(CollIdx) = 2 ^ CollIdx
        Next
        Msk(15) = -32768
    End If
    For SCUBound = 0 To UBound(SpriteColl)
        Set Spr = SpriteColl(SCUBound)
        If Spr.rDef.Template.ControlType = CONTROL_SIMPLE Then
            Exit For
        End If
    Next
    SCUBound = SCUBound - 1
        
    For CollIdx = 0 To pMap.CollDefCount - 1
        With pMap
            bDoTest = True
            With .CollDefs(CollIdx)
                If .InvUseCount <> 0 Then
                    If ((.InvFlags And eCollInvFlags.COLL_INV_REQUIRE) > 0) Xor _
                        Prj.GamePlayer.InvQuantityOwned(.InvItem) >= .InvUseCount Then
                        bDoTest = False
                    End If
                End If
                
                ClassA = .ClassA
                ClassB = .ClassB
            End With
            If (ClsMbrCount(ClassA) <= ClsMbrCount(ClassB)) And bDoTest Then
                For SprAIdx = 0 To SCUBound
                    If SprAIdx > SCUBound Then Exit For ' Number of sprites can change in loop
                    Set Spr = SpriteColl(SprAIdx)
                    SprClsMsk = Spr.rDef.Template.CollClass
                    If SprClsMsk And Msk(ClassA) Then
                        For SprBIdx = IIf(ClassA = ClassB, SprAIdx + 1, 0) To SCUBound
                            If SprBIdx > SCUBound Or SprAIdx > SCUBound Then Exit For ' Number of sprites can change in loop
                            Set Spr = SpriteColl(SprBIdx)
                            If SprAIdx <> SprBIdx Then
                                SprClsMsk = Spr.rDef.Template.CollClass
                                If SprClsMsk And Msk(ClassB) Then
                                    TestForCollision SprAIdx, SprBIdx, CollIdx
                                    If SCUBound > UBound(SpriteColl) Then SCUBound = UBound(SpriteColl)
                                End If
                            End If
                        Next
                    End If
                Next
            ElseIf bDoTest Then
                For SprBIdx = 0 To SCUBound
                    If SprBIdx > SCUBound Then Exit For
                    Set Spr = SpriteColl(SprBIdx)
                    SprClsMsk = Spr.rDef.Template.CollClass
                    If SprClsMsk And Msk(ClassB) Then
                        For SprAIdx = 0 To SCUBound
                            If SprAIdx > SCUBound Or SprBIdx > SCUBound Then Exit For
                            Set Spr = SpriteColl(SprAIdx)
                            If SprBIdx <> SprAIdx Then
                                SprClsMsk = Spr.rDef.Template.CollClass
                                If SprClsMsk And Msk(ClassA) Then
                                    TestForCollision SprAIdx, SprBIdx, CollIdx
                                    If SCUBound > UBound(SpriteColl) Then SCUBound = UBound(SpriteColl)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End With
    Next
End Sub

Public Sub TestForCollision(ByVal ClsASprIdx As Integer, ByVal ClsBSprIdx As Integer, ByVal CollDefIdx As Integer)
Attribute TestForCollision.VB_Description = "Test a specific CollisionDef given two applicable sprites.  If they are colliding, take appropriate action."
    Dim TmpD As Single
    Dim Flags As Integer
    Dim NewASpr As Sprite
    Dim NewBSpr As Sprite
    Dim SprTpl As SpriteTemplate
    Dim SprDef As SpriteDef
    Dim bCollided As Boolean
    Dim strErrMsg As String
    Dim ASpr As Sprite
    Dim BSpr As Sprite
    
    On Error GoTo TestCollError

    Flags = pMap.CollDefs(CollDefIdx).Flags

    Set ASpr = SpriteColl(ClsASprIdx)
    Set BSpr = SpriteColl(ClsBSprIdx)
    If Flags And eCollisionFlags.COLL_PLATFORM Then
        If TileRectHit(ASpr.CurTS, BSpr.CurTS, CInt(BSpr.X - ASpr.X), CInt(BSpr.Y - ASpr.Y)) Then
            Sprite(ClsASprIdx).RideOn Sprite(ClsBSprIdx)
            bCollided = True
        End If
    End If
    If (Flags And Not eCollisionFlags.COLL_PLATFORM) Or ((Flags And eCollisionFlags.COLL_PLATFORM) = 0) Then
        If pMap.Disp.TileHitTile(ASpr.CurTS, BSpr.CurTS, ASpr.CurTile, BSpr.CurTile, CInt(BSpr.X - ASpr.X), CInt(BSpr.Y - ASpr.Y)) Then
            bCollided = True
            If Flags And eCollisionFlags.COLL_ANEW Then
                Set NewASpr = ASpr.rDef.MakeInstance
                AddSprite NewASpr
            End If
            If Flags And eCollisionFlags.COLL_BNEW Then
                Set NewBSpr = BSpr.rDef.MakeInstance
                AddSprite NewBSpr
            End If
            If Flags And eCollisionFlags.COLL_SWAPVEL Then
                TmpD = ASpr.DX
                If Sgn(BSpr.X - ASpr.X) <> Sgn(BSpr.DX) Then
                    ASpr.DX = BSpr.DX
                End If
                If Sgn(ASpr.X - BSpr.X) <> Sgn(TmpD) Then
                    BSpr.DX = TmpD
                End If
                TmpD = ASpr.DY
                If Sgn(BSpr.Y - ASpr.Y) <> Sgn(BSpr.DY) Then
                    ASpr.DY = BSpr.DY
                End If
                If Sgn(ASpr.Y - BSpr.Y) <> Sgn(TmpD) Then
                    BSpr.DY = TmpD
                End If
                If Flags And eCollisionFlags.COLL_ANEW Then
                    NewASpr.X = ASpr.X
                    NewASpr.Y = ASpr.Y
                End If
                If Flags And eCollisionFlags.COLL_BNEW Then
                    NewBSpr.X = BSpr.X
                    NewBSpr.Y = BSpr.Y
                End If
            End If
            If Flags And eCollisionFlags.COLL_ASTOP Then
                Set SprDef = ASpr.rDef.Clone
                Set SprTpl = ASpr.rDef.Template.Clone
                Set ASpr.rDef = SprDef
                Set ASpr.rDef.Template = SprTpl
                ASpr.DX = 0
                ASpr.DY = 0
                ASpr.rDef.Template.ControlType = CONTROL_INERT
            End If
            If Flags And eCollisionFlags.COLL_BSTOP Then
                Set SprDef = BSpr.rDef.Clone
                Set SprTpl = BSpr.rDef.Template.Clone
                Set BSpr.rDef = SprDef
                Set BSpr.rDef.Template = SprTpl
                BSpr.DX = 0
                BSpr.DY = 0
                BSpr.rDef.Template.ControlType = CONTROL_INERT
            End If
            If Flags And eCollisionFlags.COLL_ATERM Then
                RemoveSprite ClsASprIdx
                If ClsASprIdx < ClsBSprIdx Then ClsBSprIdx = ClsBSprIdx - 1
                ClsASprIdx = -1
            End If
            If Flags And eCollisionFlags.COLL_BTERM Then
                RemoveSprite ClsBSprIdx
                ClsBSprIdx = -1
            End If
        End If
    End If
    If bCollided Then
        With pMap.CollDefs(CollDefIdx)
            If Len(.SpecialFunction) > 0 Then
                If Not Prj.GamePlayer Is Nothing Then
                    Prj.GamePlayer.ActivateFunction pMap.Specials(pMap.CollDefs(CollDefIdx).SpecialFunction)
                End If
            End If
            If Not pMap.bDisablePlayerEdit Then
                If ((.InvFlags And eCollInvFlags.COLL_INV_REQUIRE) > 0) And _
                   ((.InvFlags And eCollInvFlags.COLL_INV_REMOVE) > 0) Then
                    Prj.GamePlayer.InvQuantityOwned(.InvItem) = Prj.GamePlayer.InvQuantityOwned(.InvItem) - .InvUseCount
                End If
            End If
            If Len(.Media) Then Prj.MediaMgr.Clip(.Media).Play
        End With
        Prj.GamePlayer.TriggerCollide Name, ClsASprIdx, ClsBSprIdx, CollDefIdx
    End If
    Exit Sub

TestCollError:
    strErrMsg = Err.Description
    If Not (CurDisp Is Nothing) Then
        CurDisp.Close
        Set CurDisp = Nothing
    End If
    If Not GameHost Is Nothing Then GameHost.CheckForError
    MsgBox strErrMsg, vbExclamation
End Sub

Function TileRectHit(TileSet1 As IBMDXTileSet, TileSet2 As IBMDXTileSet, ByVal DX As Integer, ByVal DY As Integer) As Boolean
Attribute TileRectHit.VB_Description = "Determine if two sprite's rectangles overlap (used for detecting collision with a platform)."
    If TileSet1.TileWidth > DX And TileSet1.TileHeight > DY And TileSet2.TileWidth > -DX And TileSet2.TileHeight > -DY Then TileRectHit = True
End Function
