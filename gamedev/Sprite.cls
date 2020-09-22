VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a single currently active sprite instance on a layer.  Indexed and owned by layer.  Not persistent."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Sprite.cls - Sprite Instance Class
'
'======================================================================

Option Explicit

Public X As Single
Attribute X.VB_VarDescription = "Current X position of the sprite within the layer."
Public Y As Single
Attribute Y.VB_VarDescription = "Current Y position of the sprite within the layer."
Public DX As Single
Attribute DX.VB_VarDescription = "Delta-X (x velocity) of this sprite."
Public DY As Single
Attribute DY.VB_VarDescription = "Delta-Y (y velocity) of this sprite."
Public CurState As Byte ' Used as Angle for rotating sprites
Attribute CurState.VB_VarDescription = "Get/set which state the sprite is currently in."
Public CurFrame As Byte ' Index into frame array
Attribute CurFrame.VB_VarDescription = "Get/set which image (from teh current state's animation) is currently being displayed."
Public CurDelay As Byte
Attribute CurDelay.VB_VarDescription = "How many frames into the delay for the current image is the animation state of the sprite."
Public rDef As SpriteDef
Attribute rDef.VB_VarProcData = ";Relational"
Attribute rDef.VB_VarDescription = "Reference to this sprite's SpriteDef."
Public NextPointIndex As Integer
Attribute NextPointIndex.VB_VarDescription = "Get/set the index of the point (on SpriteDef's Path) toward which the sprite is moving."
Public pRideOnRef As Sprite
Attribute pRideOnRef.VB_VarDescription = "Reference to the sprite that this sprite is riding on (platform)."
Public RideRelX As Single
Attribute RideRelX.VB_VarDescription = "X position relative to a platform the sprite is riding on."
Public UserData As Variant
Attribute UserData.VB_VarDescription = "Not touched by GameDev, to be used in script or external controlling programs."

Private byAccelDir As Byte ' 1 = horizontal, 2 = vertical, 3 = both

Public Function RideOn(Platform As Sprite)
Attribute RideOn.VB_Description = "Check if sprite is going the correct direction to land on the platform, if so, ride on it."
    If pRideOnRef Is Nothing Then
        If rDef.Template.GravPow >= 10 And DY > Platform.DY Then
            If Int(Y - DY + Platform.DY + Height) <= -Int(-Platform.Y) Then
                RideRelX = X - Platform.X
                Set pRideOnRef = Platform
                DX = DX - Platform.DX
                DY = 0
            End If
        ElseIf rDef.Template.GravPow < 10 And DY < Platform.DY Then
            If -Int(-(Y - DY + Platform.DY)) > Int(Platform.Y + Platform.Height) Then
                RideRelX = X - Platform.X
                Set pRideOnRef = Platform
                DX = DX - Platform.DX
                DY = 0
            End If
        End If
    End If
End Function

Public Function ProcessAction(ByVal Actions As Integer) As Boolean ' True if accelerating
Attribute ProcessAction.VB_Description = "Set DX, DY and CurState based on passed Action flags (ie. Player.CtlActions), return True if accellerating."
    Dim Tmp As Single

    byAccelDir = 0
    If rDef.Template.StateType = STATE_36_DIRECTION Then
        CurState = CurState Mod 36
        If Actions And eActionBits.ACTION_UP Then
            DX = DX + Cos(CurState * Pi / 18)
            DY = DY - Sin(CurState * Pi / 18)
        End If
        If Actions And eActionBits.ACTION_DOWN Then
            DX = DX - Cos(CurState * Pi / 18)
            DY = DY + Sin(CurState * Pi / 18)
        End If
        If Actions And eActionBits.ACTION_LEFT Then
            CurState = (CurState + 1) Mod 36
        End If
        If Actions And eActionBits.ACTION_RIGHT Then
            CurState = (CurState + 35) Mod 36
        End If
        If (Actions And (eActionBits.ACTION_UP Or eActionBits.ACTION_DOWN)) <> 0 Then
            If (rDef.Template.Flags And FLAG_ACCELSTATES) <> 0 Then
                CurState = CurState + 36
            End If
            byAccelDir = 3
            ProcessAction = True
        End If
        Tmp = Sqr(DX * DX + DY * DY)
        If Tmp > rDef.Template.MoveSpeed Then
            ' Normalize to velocity of MoveSpeed
            DX = DX * rDef.Template.MoveSpeed / Tmp
            DY = DY * rDef.Template.MoveSpeed / Tmp
        End If
    Else
        If Actions And eActionBits.ACTION_LEFT Then
            If DX > -rDef.Template.MoveSpeed Then
                DX = DX - 1
            End If
            byAccelDir = byAccelDir Or 1
            ProcessAction = True
        End If
        If Actions And eActionBits.ACTION_RIGHT Then
            If DX < rDef.Template.MoveSpeed Then
                DX = DX + 1
            End If
            byAccelDir = byAccelDir Or 1
            ProcessAction = True
        End If
        If rDef.Template.Flags And eTemplateFlagBits.FLAG_UPNEEDSSOLID Then
            If Actions And eActionBits.ACTION_UP Then
                If (rDef.SolidTest(X, Y + Height) Or rDef.SolidTest(X + Width - 1, Y + Height)) Or (Not pRideOnRef Is Nothing) Then
                    DY = -rDef.Template.MoveSpeed
                End If
                byAccelDir = byAccelDir Or 2
                ProcessAction = True
            End If
            If Actions And eActionBits.ACTION_DOWN Then
                If (rDef.SolidTest(X, Y - 1) Or rDef.SolidTest(X + Width - 1, Y - 1)) Or (Not pRideOnRef Is Nothing) Then
                    DY = rDef.Template.MoveSpeed
                End If
                byAccelDir = byAccelDir Or 2
                ProcessAction = True
            End If
        Else
            If Actions And eActionBits.ACTION_UP Then
                If DY > -rDef.Template.MoveSpeed Then
                    DY = DY - 1
                End If
                byAccelDir = byAccelDir Or 2
                ProcessAction = True
            End If
            If Actions And eActionBits.ACTION_DOWN Then
                If DY < rDef.Template.MoveSpeed Then
                    DY = DY + 1
                End If
                byAccelDir = byAccelDir Or 2
                ProcessAction = True
            End If
        End If
    End If
End Function

Public Sub ReactToSolid()
Attribute ReactToSolid.VB_Description = "Alter DX and DY according to map solidity via SolidDef in Sprite's template."
    Dim W As Integer, H As Integer
    Dim bConst As Boolean
        
    W = Width
    H = Height

    bConst = Not rDef.SolidTest(X + DX + W - 1, Y + H - 1) And Not rDef.SolidTest(X + DX, Y + H - 1)
    Do While (DY > 0) And _
            ((rDef.SolidTest(X + DX, Y + DY + H - 1) Or _
            rDef.SolidTest(X + DX + W - 1, Y + DY + H - 1)) And _
            (rDef.SolidTest(X + W - 1, Y + DY + H - 1) Or rDef.SolidTest(X, Y + DY + H - 1) Or bConst))
        DY = DY - 1
        If DY <= 0 Then DY = 0: Exit Do
    Loop

    bConst = Not rDef.SolidTest(X + DX + W - 1, Y) And Not rDef.SolidTest(X + DX, Y)
    Do While (DY < 0) And _
            ((rDef.SolidTest(X + DX, Y + DY) Or _
            rDef.SolidTest(X + DX + W - 1, Y + DY)) And _
            (rDef.SolidTest(X + W - 1, Y + DY) Or rDef.SolidTest(X, Y + DY) Or bConst))
        DY = DY + 1
        If DY >= 0 Then DY = 0: Exit Do
    Loop
    
    Do While (DX > 0) And _
            (rDef.SolidTest(X + DX + W - 1, Y + DY) Or _
            rDef.SolidTest(X + DX + W - 1, Y + H - 1 + DY))
         
        ' Check uphill (rightward) stuff
        If (Not rDef.SolidTest(X + DX + W - 1, Y + H - 1.5 - DX)) And _
            (Not rDef.SolidTest(X + DX + W - 1, Y - DX)) And _
            (Not rDef.SolidTest(X + DX, Y - DX)) Then
            DY = -DX - 0.5
            Exit Do
        ' Check down-ceiling (rightward) stuff
        ElseIf (Not rDef.SolidTest(X + DX + W - 1, Y + H - 1 + DX)) And _
               (Not rDef.SolidTest(X + DX + W - 1, Y + DX)) And _
               (Not rDef.SolidTest(X + DX, Y + H - 1 + DX)) Then
            DY = DX
            Exit Do
        Else
            DX = DX - 1
            If DX < 0 Then DX = 0: Exit Do
        End If
   Loop
   
   Do While (DX < 0) And _
            (rDef.SolidTest(X + DX, Y + DY) Or _
            rDef.SolidTest(X + DX, Y + H - 1 + DY))
            
        ' Check uphill (leftward) stuff
        If rDef.SolidTest(X + DX, Y + DY + H - 1) And _
            (Not rDef.SolidTest(X + DX, Y + H - 1.5 + DX) And _
            (Not rDef.SolidTest(X + DX, Y + DX)) And _
            (Not rDef.SolidTest(X + W - 1 + DX, Y + DX))) Then
            DY = DX - 0.5
            Exit Do
        ' Check down-ceiling (leftward) stuff
        ElseIf rDef.SolidTest(X + DX, Y + DY) And _
            (Not rDef.SolidTest(X + DX, Y + H - 1 - DX) And _
            (Not rDef.SolidTest(X + DX, Y - DX)) And _
            (Not rDef.SolidTest(X + W - 1 + DX, Y + H - 1 - DX))) Then
            DY = -DX
        Else
           DX = DX + 1
           If DX > 0 Then DX = 0: Exit Do
        End If
    Loop
    
End Sub

Public Property Get Width() As Integer
Attribute Width.VB_Description = "The width of the sprite in its current state."
    If CurTSDef Is Nothing Then
        Width = 0
    Else
        Width = CurTSDef.TileWidth
    End If
End Property

Public Property Get Height() As Integer
Attribute Height.VB_Description = "Return the height of the sprite in its current state."
    If CurTSDef Is Nothing Then
        Height = 0
    Else
        Height = CurTSDef.TileHeight
    End If
End Property

Public Property Get CurTile() As Integer
Attribute CurTile.VB_Description = "Returns the tile index within the sprite's tileset that is used to draw the sprite in its current state."
    CurTile = rDef.Template.StateFrame(CurState, CurFrame)
End Property

Public Property Get CurTS() As IBMDXTileSet
Attribute CurTS.VB_Description = "Returns the tileset used to draw the sprite in its crrent state."
    Dim TSD As TileSetDef
    
    Set TSD = CurTSDef
    If Not (TSD Is Nothing) Then
        Set CurTS = TSD.LoadedInstance
    End If
End Property

Public Property Get CurTSDef() As TileSetDef
Attribute CurTSDef.VB_Description = "Returns the TileSetDef used to draw the sprite in its current state."
    Set CurTSDef = rDef.Template.StateTilesetDef(CurState)
End Property

Public Sub AdvanceFrame(ByVal Distance As Integer)
Attribute AdvanceFrame.VB_Description = "Advance the image for this sprite by a specified number of frames."
    If rDef.Template.StateFrameCount(CurState) Then
        CurFrame = (CurFrame + Int((Distance + CurDelay) / (rDef.Template.AnimSpeed + 1))) Mod rDef.Template.StateFrameCount(CurState)
        CurDelay = (CurDelay + Distance) Mod (rDef.Template.AnimSpeed + 1)
    End If
End Sub

Public Sub ResetFrame()
Attribute ResetFrame.VB_Description = "Restart sprite's animation (without changing state)."
    CurFrame = 0
    CurDelay = 0
End Sub

Private Sub ReactToPlatform()
    Dim WMe As Integer
    Dim WRide As Integer
    
    WMe = Width
    WRide = pRideOnRef.Width
    
    RideRelX = RideRelX + DX
    X = pRideOnRef.X + RideRelX
    If rDef.Template.GravPow >= 10 Then
        Y = pRideOnRef.Y - Height
        If DY < 0 Then
            Y = Y + DY
            DX = DX + pRideOnRef.DX
            Set pRideOnRef = Nothing
        End If
    Else
        Y = pRideOnRef.Y + pRideOnRef.Height
        If DY > 0 Then
            Y = Y + DY
            DX = DX + pRideOnRef.DX
            Set pRideOnRef = Nothing
        End If
    End If
    If Not pRideOnRef Is Nothing Then
        If X + WMe <= pRideOnRef.X Or pRideOnRef.X + WRide <= X Then
            DX = DX + pRideOnRef.DX
            DY = pRideOnRef.DY
            Set pRideOnRef = Nothing
        End If
    End If
End Sub

Public Sub Advance()
Attribute Advance.VB_Description = "Move this sprite to its next position based on it's SpriteDef and SpriteTemplate."
    Dim Ang As Single ' Heading to next point (adjusted to be desired heading)
    Dim Dis As Single ' Distance to next point
    Dim CurAng As Single ' Current heading (may be different than direction of travel)
    Dim CurVel As Single ' Current velocity
    Dim VelAng As Single ' Angle of the sprite's velocity
    Dim Actions As Integer
    Dim NormalX As Single, NormalY As Single
    Dim NormalDot As Single
    Dim I As Integer
    Dim CloseIndex As Integer
    Dim W As Integer, H As Integer
    Dim bLost As Boolean, bLost2 As Boolean
    Dim bAccel As Boolean
   
    With rDef
        
        DY = DY + CSng(.Template.GravPow - 10) / 10
        
        If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Or .Template.ControlType = CONTROL_FOLLOWRIGHTWALL Then
            X = Int(X + 0.5): Y = Int(Y + 0.5): DX = Int(DX + 0.5): DY = Int(DY + 0.5)
            
            If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Then
                GetSuggestedLeftWallDeltas X, Y, DX, DY, Dis, CurVel, bLost
            Else
                GetSuggestedRightWallDeltas X, Y, DX, DY, Dis, CurVel, bLost
            End If

            If Dis = 0 And CurVel = 0 Then
                If DX = 0 And DY = 0 Then
                    If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Then
                        DX = -1
                    Else
                        DX = 1
                    End If
                Else
                    For I = 1 To .Template.MoveSpeed
                        If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Then
                            GetSuggestedLeftWallDeltas X + Sgn(DX) * I, Y + Sgn(DY) * I, DX, DY, NormalX, NormalY, bLost2
                        Else
                            GetSuggestedRightWallDeltas X + Sgn(DX) * I, Y + Sgn(DY) * I, DX, DY, NormalX, NormalY, bLost2
                        End If
                        If NormalX <> Dis Or NormalY <> CurVel Or bLost2 <> bLost Then Exit For
                    Next
                    If I > .Template.MoveSpeed Then I = .Template.MoveSpeed
                    DX = Sgn(DX) * I
                    DY = Sgn(DY) * I
                End If
            Else
                For I = 1 To .Template.MoveSpeed
                    If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Then
                        GetSuggestedLeftWallDeltas X + Sgn(Dis) * I, Y + Sgn(CurVel) * I, Dis, CurVel, NormalX, NormalY, bLost2
                    Else
                        GetSuggestedRightWallDeltas X + Sgn(Dis) * I, Y + Sgn(CurVel) * I, Dis, CurVel, NormalX, NormalY, bLost2
                    End If
                    If NormalX <> Dis Or NormalY <> CurVel Or bLost2 <> bLost Then Exit For
                Next
                If I > .Template.MoveSpeed Then I = .Template.MoveSpeed
                DX = Sgn(Dis) * I
                DY = Sgn(CurVel) * I
            End If
            Dis = DX
            CurVel = DY
            ReactToSolid
            I = 0
            Do
                I = I + 1
                If DX = 0 And DY = 0 Or DX <> Int(DX) Or DY <> Int(DY) Then
                    If .Template.ControlType = CONTROL_FOLLOWLEFTWALL Then
                        RotateMotion Dis, CurVel, True
                    Else
                        RotateMotion Dis, CurVel, False
                    End If
                    DX = Dis
                    DY = CurVel
                End If
                ReactToSolid
            Loop Until I > 4 Or ((DX <> 0 Or DY <> 0) And DX = Int(DX) And DY = Int(DY))
        End If

        If .Template.StateType = STATE_36_DIRECTION Then
            Select Case .Template.ControlType
            Case CONTROL_APPROX_PATH
                CurAng = CurState * Pi / 18
                RectToPolar DX, DY, VelAng, CurVel
                
                NormalX = .rPath.PointY(NextPointIndex) - Y
                NormalY = .rPath.PointX(NextPointIndex) - X
                NormalDot = DotProduct(NormalX, -NormalY, DX, DY)
                If NormalDot < 0 Then
                    NormalX = -NormalX * NormalDot / 2
                    NormalY = NormalY * NormalDot / 2
                Else
                    NormalDot = DotProduct(-NormalX, NormalY, DX, DY)
                    NormalX = NormalX * NormalDot / 2
                    NormalY = -NormalY * NormalDot / 2
                End If

                RectToPolar .rPath.PointX(NextPointIndex) - X + NormalX, .rPath.PointY(NextPointIndex) - Y + NormalY, Ang, Dis
                If Abs(Ang - CurAng) > Pi Then
                    If Ang < CurAng Then Ang = Ang + Pi * 2 Else Ang = Ang - Pi * 2
                End If
                If Abs(Ang - CurAng) > Pi / 36 Then
                    If Ang < CurAng Then Actions = eActionBits.ACTION_RIGHT Else Actions = eActionBits.ACTION_LEFT
                    If Dis <= (CurVel * (CurVel + 1)) / 2 Then
                        Actions = Actions Or eActionBits.ACTION_DOWN
                    End If
                Else
                    If Dis - 12 > (CurVel * (CurVel + 1)) / 2 Then
                        Actions = Actions Or eActionBits.ACTION_UP
                    Else
                        Actions = Actions Or eActionBits.ACTION_DOWN
                    End If
                End If
                                                                
                If Dis < 12 Then
                    NextPointIndex = NextPointIndex + 1
                    If NextPointIndex >= .rPath.PointCount Then NextPointIndex = 0
                End If
                bAccel = ProcessAction(Actions)
                If pRideOnRef Is Nothing Then
                    ReactToSolid
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_EXACT_PATH
                CurAng = CurState * Pi / 18
                
                CurState = CurState Mod 36
                RectToPolar .rPath.PointX(NextPointIndex) - X, .rPath.PointY(NextPointIndex) - Y, Ang, Dis
                If Abs(Ang - CurAng) > Pi Then
                    If Ang < CurAng Then Ang = Ang + Pi * 2 Else Ang = Ang - Pi * 2
                End If
                If Abs(Ang - CurAng) > Pi / 36 Then
                    If Ang < CurAng Then CurState = (CurState + 35) Mod 36 Else CurState = (CurState + 1) Mod 36
                    DX = 0
                    DY = 0
                Else
                    NormalX = .rPath.PointX(NextPointIndex) - X
                    NormalY = .rPath.PointY(NextPointIndex) - Y
                    If Dis - 12 > (.Template.MoveSpeed * (.Template.MoveSpeed + 1)) / 2 Then
                        DX = NormalX * .Template.MoveSpeed / Dis
                        DY = NormalY * .Template.MoveSpeed / Dis
                        If .Template.Flags And FLAG_ACCELSTATES Then CurState = CurState + 36
                        bAccel = True
                    ElseIf Dis > 0 Then
                        DX = NormalX / Sqr(Dis)
                        DY = NormalY / Sqr(Dis)
                    Else
                        DX = 0
                        DY = 0
                    End If
                End If
                                                                
                If Dis < 2 Then
                    X = .rPath.PointX(NextPointIndex)
                    Y = .rPath.PointY(NextPointIndex)
                    NextPointIndex = NextPointIndex + 1
                    If NextPointIndex >= .rPath.PointCount Then NextPointIndex = 0
                    DX = 0
                    DY = 0
                End If
                If pRideOnRef Is Nothing Then
                    ReactToSolid
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_INPUT
                bAccel = ProcessAction(Prj.GamePlayer.CtlActions)
                If pRideOnRef Is Nothing Then
                    ReactToSolid
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_FOLLOWNEAR, CONTROL_FOLLOWPLAYER
                CloseIndex = -1
                If .Template.ControlType = CONTROL_FOLLOWNEAR Then
                    Dis = 1E+09!
                    For I = 0 To .rLayer.SpriteCount - 1
                        If Not (.rLayer.Sprite(I) Is Me) Then
                            With .rLayer.Sprite(I)
                                CurVel = Sqr((.X - X) * (.X - X) + (.Y - Y) * (.Y - Y))
                                If CurVel < Dis Then
                                    Dis = CurVel
                                    CloseIndex = I
                                End If
                            End With
                        End If
                    Next
                Else
                    For I = 0 To .rLayer.SpriteCount - 1
                        If Not (.rLayer.Sprite(I) Is Me) Then
                            If .rLayer.Sprite(I) Is Prj.GamePlayer.PlayerSprite Then
                                CloseIndex = I
                                Exit For
                            End If
                        End If
                    Next
                End If
                CurAng = CurState * Pi / 18
                If CloseIndex >= 0 Then
                    With .rLayer.Sprite(CloseIndex)
                        RectToPolar .X - X, .Y - Y, Ang, Dis
                    End With
                    
                    If Abs(Ang - CurAng) > Pi Then
                        If Ang < CurAng Then Ang = Ang + Pi * 2 Else Ang = Ang - Pi * 2
                    End If
                    If Abs(Ang - CurAng) > Pi / 36 Then
                        If Ang < CurAng Then Actions = eActionBits.ACTION_RIGHT Else Actions = eActionBits.ACTION_LEFT
                    Else
                        Actions = Actions Or eActionBits.ACTION_UP
                    End If
                    bAccel = ProcessAction(Actions)
                    If pRideOnRef Is Nothing Then
                        ReactToSolid
                        X = X + DX
                        Y = Y + DY
                    Else
                        ReactToPlatform
                    End If
                    AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
                End If
            Case CONTROL_FOLLOWFLATFLOOR
                CurState = CurState Mod 36
                If CurState = 0 Or CurState = 18 Then
                    Actions = eActionBits.ACTION_UP
                    bAccel = ProcessAction(Actions)
                    If CurState = 18 Then
                        If rDef.SolidTest(X + DX, Y + Height - 1) Then
                            Actions = eActionBits.ACTION_LEFT
                            bAccel = ProcessAction(Actions)
                        End If
                    Else
                        If rDef.SolidTest(X + Width + DX, Y + Height - 1) Then
                            Actions = eActionBits.ACTION_LEFT
                            bAccel = ProcessAction(Actions)
                        End If
                    End If
                    ReactToSolid
                Else
                    Actions = eActionBits.ACTION_LEFT
                    bAccel = ProcessAction(Actions)
                    ReactToSolid
                End If
                
                If pRideOnRef Is Nothing Then
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Abs(DX)
            Case CONTROL_FOLLOWSLOPEFLOOR
                CurState = CurState Mod 36
                If CurState = 0 Or CurState = 18 Then
                    CurVel = Abs(DX)
                    Actions = eActionBits.ACTION_UP
                    bAccel = ProcessAction(Actions)
                    ReactToSolid
                    If DX = 0 Then CurState = CurState + 1
                Else
                    Actions = eActionBits.ACTION_LEFT
                    bAccel = ProcessAction(Actions)
                    ReactToSolid
                End If
                If pRideOnRef Is Nothing Then
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Abs(DX)
            Case CONTROL_FOLLOWFLATWALL
                W = Width
                H = Height
                CurState = CurState Mod 36
                If .SolidTest(X + W, Y) Then
                    If .SolidTest(X + W, Y + H - 1) Then
                        If CurState = 9 Then
                            If .SolidTest(X, Y - 1) Or .SolidTest(X + W - 1, Y - 1) Then
                                Actions = eActionBits.ACTION_LEFT
                            Else
                                Actions = eActionBits.ACTION_UP
                            End If
                        ElseIf CurState = 27 Then
                            If .SolidTest(X, Y + H) Or .SolidTest(X + W - 1, Y + H) Then
                                Actions = eActionBits.ACTION_RIGHT
                            Else
                                Actions = eActionBits.ACTION_UP
                            End If
                        Else
                            If .SolidTest(X, Y + H) Or .SolidTest(X + W - 1, Y + H) Then
                                Actions = eActionBits.ACTION_RIGHT
                            Else
                                Actions = eActionBits.ACTION_LEFT
                            End If
                        End If
                    Else
                        If .SolidTest(X + W - 1, Y - 1) Then
                            If CurState = 27 Then
                                Actions = eActionBits.ACTION_UP
                            Else
                                Actions = eActionBits.ACTION_RIGHT
                            End If
                        Else
                            If CurState = 9 Then
                                Actions = eActionBits.ACTION_UP
                            Else
                                Actions = eActionBits.ACTION_RIGHT
                                If CurState > 18 And DY > 1 Then Actions = Actions Or eActionBits.ACTION_DOWN
                            End If
                        End If
                    End If
                Else
                    If .SolidTest(X + W, Y + H - 1) Then
                        If CurState = 27 Then
                            Actions = eActionBits.ACTION_UP
                        Else
                            Actions = eActionBits.ACTION_LEFT
                            If CurState < 18 And DY < -1 Then Actions = Actions Or eActionBits.ACTION_DOWN
                        End If
                    Else
                        If .SolidTest(X - 1, Y) Then
                            If .SolidTest(X - 1, Y + H - 1) Then
                                If CurState = 9 Then
                                    If .SolidTest(X, Y - 1) Or .SolidTest(X + W - 1, Y - 1) Then
                                        Actions = eActionBits.ACTION_RIGHT
                                    Else
                                        Actions = eActionBits.ACTION_UP
                                    End If
                                ElseIf CurState = 27 Then
                                    If .SolidTest(X, Y + H) Or .SolidTest(X + W - 1, Y + H) Then
                                        Actions = eActionBits.ACTION_LEFT
                                    Else
                                        Actions = eActionBits.ACTION_UP
                                    End If
                                Else
                                    If .SolidTest(X, Y + H) Or .SolidTest(X + W - 1, Y + H) Then
                                        Actions = eActionBits.ACTION_LEFT
                                    Else
                                        Actions = eActionBits.ACTION_RIGHT
                                    End If
                                End If
                            Else
                                If CurState = 9 Then
                                    Actions = eActionBits.ACTION_UP
                                Else
                                    Actions = eActionBits.ACTION_LEFT
                                    If CurState >= 27 And DY > 1 Then Actions = Actions Or eActionBits.ACTION_DOWN
                                End If
                            End If
                        Else
                            If .SolidTest(X - 1, Y + H - 1) Then
                                If CurState = 27 Then
                                    Actions = eActionBits.ACTION_UP
                                Else
                                    Actions = eActionBits.ACTION_RIGHT
                                    If CurState < 18 Then Actions = Actions Or eActionBits.ACTION_DOWN
                                End If
                            Else
                                If .SolidTest(X + W * 2, Y + H / 2) Then
                                    If CurState < 18 Then
                                        If CurState > 0 Then Actions = eActionBits.ACTION_RIGHT Or eActionBits.ACTION_DOWN
                                        If CurState < 9 Then Actions = Actions Or eActionBits.ACTION_UP
                                    Else
                                        Actions = eActionBits.ACTION_LEFT
                                        If CurState > 27 Then Actions = Actions Or eActionBits.ACTION_UP
                                    End If
                                Else
                                    If CurState > 18 Then
                                        Actions = eActionBits.ACTION_RIGHT
                                        If CurState < 27 Then Actions = Actions Or eActionBits.ACTION_UP
                                    Else
                                        If CurState < 18 Then Actions = eActionBits.ACTION_LEFT Or eActionBits.ACTION_DOWN
                                        If CurState > 9 Then Actions = Actions Or eActionBits.ACTION_UP
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                bAccel = ProcessAction(Actions)
                If pRideOnRef Is Nothing Then
                    ReactToSolid
                    X = X + DX
                    Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_FOLLOWLEFTWALL, CONTROL_FOLLOWRIGHTWALL
                RectToPolar DX, DY, Ang, Dis
                CurState = CurState Mod 36
                CurAng = Int(((Ang * (180 / Pi)) Mod 360) / 10)
                If CurAng <> CurState Then
                    If CurAng - CurState < 0 Then
                        If CurAng + 36 - CurState < 18 Then
                            CurState = (CurState + 1) Mod 36
                        Else
                            CurState = (CurState + 35) Mod 36
                        End If
                    Else
                        If CurAng - CurState < 18 Then
                            CurState = (CurState + 1) Mod 36
                        Else
                            CurState = (CurState + 35) Mod 36
                        End If
                    End If
                Else
                    If pRideOnRef Is Nothing Then
                        X = X + DX
                        Y = Y + DY
                    Else
                        ReactToPlatform
                    End If
                    If .Template.Flags And FLAG_ACCELSTATES Then CurState = CurState + 36
                    bAccel = True
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_SLOPEFLOORJUMP
                W = Width
                H = Height
                If DX > 0 Then
                    DX = .Template.MoveSpeed
                Else
                    If .SolidTest(X - 1, Y) Or (.SolidTest(X - 1, Y + H - 1) And (.SolidTest(X - 1, Y - 1) Or .SolidTest(X + W - 2, Y - 1))) Then
                        DX = .Template.MoveSpeed
                    Else
                        DX = -.Template.MoveSpeed
                    End If
                End If
                If DX < 0 Then
                    If Not (.SolidTest(X, Y + H + W - 1)) And .SolidTest(X + W, Y + H) Then DY = -.Template.MoveSpeed
                Else
                    If Not (.SolidTest(X + W - 1, Y + H + W - 1)) And .SolidTest(X, Y + H) Then DY = -.Template.MoveSpeed
                End If
                ReactToSolid
                RectToPolar DX, DY, Ang, Dis
                CurState = (Int(Ang * 180 / Pi / 10) Mod 36) + IIf(DX Or DY, 36, 0)
                bAccel = True
                If pRideOnRef Is Nothing Then
                    X = X + DX: Y = Y + DY
                Else
                    ReactToPlatform
                End If
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            Case CONTROL_INERT
                ReactToSolid
                X = X + DX: Y = Y + DY
                AdvanceFrame Int(Sqr(DX * DX + DY * DY) + 0.9)
            End Select
        Else
' **** NON-ROTATIONAL SPRITES ****
            Select Case .Template.ControlType
            Case CONTROL_APPROX_PATH
                NormalX = .rPath.PointX(NextPointIndex) - X
                NormalY = .rPath.PointY(NextPointIndex) - Y
                Dis = Sqr(NormalX * NormalX + NormalY * NormalY)
                If NormalX < DX Then
                    Actions = eActionBits.ACTION_LEFT
                ElseIf NormalX > DX Then
                    Actions = eActionBits.ACTION_RIGHT
                End If
                If NormalY < DY Then
                    Actions = Actions Or eActionBits.ACTION_UP
                ElseIf NormalY > DY Then
                    Actions = Actions Or eActionBits.ACTION_DOWN
                End If

                If Dis < 12 Then
                    NextPointIndex = NextPointIndex + 1
                    If NextPointIndex >= .rPath.PointCount Then NextPointIndex = 0
                End If
                
                bAccel = ProcessAction(Actions)
                ReactToSolid
            Case CONTROL_EXACT_PATH
                Dis = Abs((.rPath.PointX(NextPointIndex) - X)) + Abs((.rPath.PointY(NextPointIndex) - Y))
                If Dis > 0 Then
                    DX = (.rPath.PointX(NextPointIndex) - X) * .Template.MoveSpeed / Dis
                    DY = (.rPath.PointY(NextPointIndex) - Y) * .Template.MoveSpeed / Dis
                    bAccel = True
                End If
                Dis = Abs(.rPath.PointX(NextPointIndex) - X)
                If Abs(.rPath.PointY(NextPointIndex) - Y) > Dis Then
                    Dis = Abs(.rPath.PointY(NextPointIndex) - Y)
                End If
                If Dis <= .Template.MoveSpeed Then
                    X = .rPath.PointX(NextPointIndex)
                    Y = .rPath.PointY(NextPointIndex)
                    NextPointIndex = NextPointIndex + 1
                    If NextPointIndex >= .rPath.PointCount Then NextPointIndex = 0
                    DX = 0
                    DY = 0
                End If
                ReactToSolid
            Case CONTROL_INPUT
                bAccel = ProcessAction(Prj.GamePlayer.CtlActions)
                ReactToSolid
            Case CONTROL_FOLLOWNEAR, CONTROL_FOLLOWPLAYER
                CloseIndex = -1
                If .Template.ControlType = CONTROL_FOLLOWNEAR Then
                    Dis = 1E+09!
                    For I = 0 To .rLayer.SpriteCount - 1
                        If Not (.rLayer.Sprite(I) Is Me) Then
                            With .rLayer.Sprite(I)
                                CurVel = Sqr((.X - X) * (.X - X) + (.Y - Y) * (.Y - Y))
                                If CurVel < Dis Then
                                    Dis = CurVel
                                    CloseIndex = I
                                End If
                            End With
                        End If
                    Next
                Else
                    For I = 0 To .rLayer.SpriteCount - 1
                        If Not (.rLayer.Sprite(I) Is Me) Then
                            If .rLayer.Sprite(I) Is Prj.GamePlayer.PlayerSprite Then
                                CloseIndex = I
                                Exit For
                            End If
                        End If
                    Next
                End If
                If CloseIndex >= 0 Then
                    Actions = 0
                    With .rLayer.Sprite(CloseIndex)
                        If .X < X Then Actions = Actions Or eActionBits.ACTION_LEFT
                        If .X > X Then Actions = Actions Or eActionBits.ACTION_RIGHT
                        If .Y < Y Then Actions = Actions Or eActionBits.ACTION_UP
                        If .Y > Y Then Actions = Actions Or eActionBits.ACTION_DOWN
                    End With
                    bAccel = ProcessAction(Actions)
                    ReactToSolid
                End If
            Case CONTROL_FOLLOWFLATFLOOR
                H = Height
                W = Width
                If DX < 0 Then
                    If .SolidTest(X - (Int(DX) * Int(DX) + Abs(DX)) / 2, Y + H) And Not .SolidTest(X - (Int(DX) * Int(DX) + Abs(DX)) / 2 - 1, Y + H - 1) Then
                        Actions = eActionBits.ACTION_LEFT
                    Else
                        Actions = eActionBits.ACTION_RIGHT
                    End If
                Else
                    If .SolidTest(X + W - 1 + (Int(-DX) * Int(-DX) + Abs(DX)) / 2, Y + H) And Not .SolidTest(X + W + (Int(-DX) * Int(-DX) + Abs(DX)) / 2, Y + H - 1) Then
                        Actions = eActionBits.ACTION_RIGHT
                    Else
                        Actions = eActionBits.ACTION_LEFT
                    End If
                End If
                bAccel = ProcessAction(Actions)
                ReactToSolid
            Case CONTROL_FOLLOWSLOPEFLOOR, CONTROL_SLOPEFLOORJUMP
                W = Width
                H = Height
                If DX > 0 Then
                    Actions = eActionBits.ACTION_RIGHT
                Else
                    If .SolidTest(X - 1, Y) Or (.SolidTest(X - 1, Y + H - 1) And (.SolidTest(X - 1, Y - 1) Or .SolidTest(X + W - 2, Y - 1))) Then
                        Actions = eActionBits.ACTION_RIGHT
                    Else
                        Actions = eActionBits.ACTION_LEFT
                    End If
                End If
                If .Template.ControlType = CONTROL_SLOPEFLOORJUMP Then
                    If Actions And eActionBits.ACTION_LEFT Then
                        If Not (.SolidTest(X, Y + H + W - 1)) Then Actions = Actions Or eActionBits.ACTION_UP
                    Else
                        If Not (.SolidTest(X + W - 1, Y + H + W - 1)) Then Actions = Actions Or eActionBits.ACTION_UP
                    End If
                End If
                bAccel = ProcessAction(Actions)
                ReactToSolid
            Case CONTROL_FOLLOWFLATWALL
                W = Width
                H = Height
                If .SolidTest(X + W, Y) Or .SolidTest(X + W, Y + H - 1) Or .SolidTest(X - 1, Y) Or .SolidTest(X - 1, Y + H - 1) Then
                    If DY < 0 Then
                        Dis = -Int(DY)
                        Dis = (Dis * Dis + Dis) / 2
                        If .SolidTest(X - 1, Y - Dis) Or .SolidTest(X + W, Y - Dis) Then
                            Actions = eActionBits.ACTION_UP
                        Else
                            Actions = eActionBits.ACTION_DOWN
                        End If
                    Else
                        If .SolidTest(X, Y + H) Or .SolidTest(X + W - 1, Y + H) Then
                            Actions = eActionBits.ACTION_UP
                        Else
                            Dis = -Int(-DY)
                            Dis = (Dis * Dis + Dis) / 2 - 1
                            If .SolidTest(X - 1, Y + H + Dis) Or .SolidTest(X + W, Y + H + Dis) Then
                                Actions = eActionBits.ACTION_DOWN
                            Else
                                Actions = eActionBits.ACTION_UP
                            End If
                        End If
                    End If
                Else
                    For Dis = W / 2 To W * 3 Step W / 2
                        If .SolidTest(X - Dis, Y) Or .SolidTest(X - Dis, Y + H - 1) Then
                            Actions = eActionBits.ACTION_LEFT
                            Exit For
                        ElseIf .SolidTest(X + W - 1 + Dis, Y) Or .SolidTest(X + W - 1 + Dis, Y + H - 1) Then
                            Actions = eActionBits.ACTION_RIGHT
                            Exit For
                        End If
                    Next
                    If Actions <> eActionBits.ACTION_LEFT Then Actions = eActionBits.ACTION_RIGHT
                End If
                bAccel = ProcessAction(Actions)
                ReactToSolid
            Case CONTROL_FOLLOWLEFTWALL, CONTROL_FOLLOWRIGHTWALL
                ' Handled at beginning of Sub
                bAccel = True
            Case CONTROL_INERT
                ReactToSolid
            End Select
            If pRideOnRef Is Nothing Then
                Dis = X
                CurVel = Y
                X = X + DX
                Y = Y + DY
            Else
                Dis = RideRelX
                ReactToPlatform
            End If
            If .Template.StateType = STATE_LEFT_RIGHT Then
                If DX < 0 Then CurState = 0 Else If DX > 0 Then CurState = 1
                If bAccel And .Template.Flags And FLAG_ACCELSTATES And CurState < 2 Then CurState = CurState + 2
                If pRideOnRef Is Nothing Then
                    AdvanceFrame Abs(Int(X) - Int(Dis))
                Else
                    AdvanceFrame Abs(Int(RideRelX) - Int(Dis))
                End If
            Else
                If .Template.StateType = STATE_8_DIRECTION Then
                    CurState = CurState Mod 8
                    If Abs(DX) < Abs(DY / 2) Then ' Vertical movement only
                        If DY < 0 Then
                            CurState = 0 ' Up
                        ElseIf DY > 0 Then
                            CurState = 4 ' Down
                        End If
                    ElseIf Abs(DY) < Abs(DX / 2) Then ' Horizontal movement only
                        If DX < 0 Then
                            CurState = 6 ' Left
                        ElseIf DX > 0 Then
                            CurState = 2 ' Right
                        End If
                    Else ' Horizontal and vertical movement
                        If DX < 0 Then
                            If DY < 0 Then
                                CurState = 7 ' Up left
                            ElseIf DY > 0 Then
                                CurState = 5 ' Down left
                            End If
                        ElseIf DX > 0 Then
                            If DY < 0 Then
                                CurState = 1 ' Up right
                            ElseIf DY > 0 Then
                                CurState = 3 ' Down right
                            End If
                        End If
                    End If
                    If bAccel And .Template.Flags And FLAG_ACCELSTATES Then CurState = CurState + 8
                Else
                    If .Template.Flags And FLAG_ACCELSTATES Then
                        If bAccel Then
                            CurState = (CurState Mod (.Template.StateCount \ 2)) + .Template.StateCount \ 2
                        Else
                            CurState = (CurState Mod (.Template.StateCount \ 2))
                        End If
                    End If
                End If
                If pRideOnRef Is Nothing Then
                    If Abs(Int(X) - Int(Dis)) >= Abs(Int(Y) - Int(CurVel)) Then
                        Dis = Abs(Int(X) - Int(Dis))
                    Else
                        Dis = Abs(Int(Y) - Int(CurVel))
                    End If
                    AdvanceFrame Dis
                Else
                    AdvanceFrame Abs(Int(RideRelX) - Int(Dis))
                End If
            End If
        End If
        
        If (.Template.Inertia < 100) Then
            If (byAccelDir And 1) = 0 Then
                DX = DX * .Template.Inertia / 100
            End If
            If .Template.GravPow = 10 Then
                If (byAccelDir And 2) = 0 Then
                    DY = DY * .Template.Inertia / 100
                End If
                If Abs(DY) < 0.05 Then DY = 0
            End If
            If Abs(DX) < 0.05 Then DX = 0
        End If

    End With
    
End Sub

Private Sub GetSuggestedLeftWallDeltas(ByVal CurX As Single, ByVal CurY As Single, ByVal CurDX As Single, ByVal CurDY As Single, SuggDX As Single, SuggDY As Single, bLost As Boolean)
    Dim W As Integer, H As Integer
    Dim BSol As Single, TSol As Single, OSol As Single
    Dim LSol As Single
    
    SuggDX = 0
    SuggDY = 0
    W = Width
    H = Height

    With rDef
        If CurDX = 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX - 1, CurY + H)
            LSol = .SolidTest(CurX - 1, CurY + H - 2)
            TSol = .SolidTest(CurX - 1, CurY)
            OSol = .SolidTest(CurX, CurY - 1)
        ElseIf CurDX > 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX - 1, CurY)
            LSol = .SolidTest(CurX, CurY - 1)
            TSol = LSol
            OSol = .SolidTest(CurX + W - 1, CurY - 1)
        ElseIf CurDX > 0 And CurDY = 0 Then
            BSol = .SolidTest(CurX - 1, CurY - 1)
            LSol = .SolidTest(CurX + 1, CurY - 1)
            TSol = .SolidTest(CurX + W - 1, CurY - 1)
            OSol = .SolidTest(CurX + W, CurY)
        ElseIf CurDX > 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX + W - 1, CurY - 1)
            LSol = .SolidTest(CurX + W, CurY)
            TSol = LSol
            OSol = .SolidTest(CurX + W, CurY + H - 1)
        ElseIf CurDX = 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX + W, CurY - 1)
            LSol = .SolidTest(CurX + W, CurY + 1)
            TSol = .SolidTest(CurX + W, CurY + H - 1)
            OSol = .SolidTest(CurX + W - 1, CurY + H)
        ElseIf CurDX < 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX + W, CurY + H - 1)
            LSol = .SolidTest(CurX + W - 1, CurY + H)
            TSol = LSol
            OSol = .SolidTest(CurX, CurY + H)
        ElseIf CurDX < 0 And CurDY = 0 Then
            BSol = .SolidTest(CurX + W, CurY + H)
            LSol = .SolidTest(CurX + W - 2, CurY + H)
            TSol = .SolidTest(CurX, CurY + H)
            OSol = .SolidTest(CurX - 1, CurY + H - 1)
        ElseIf CurDX < 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX, CurY + H)
            LSol = .SolidTest(CurX - 1, CurY + H - 1)
            TSol = LSol
            OSol = .SolidTest(CurX - 1, CurY)
        Else ' cury=0 and curx=0
            With rDef
                If .SolidTest(CurX, CurY - 1) Then SuggDX = SuggDX + 1.5
                If .SolidTest(CurX + W - 1, CurY - 1) Then SuggDX = SuggDX + 1
                If .SolidTest(CurX + W, CurY) Then SuggDY = SuggDY + 1.5
                If .SolidTest(CurX + W, CurY + H - 1) Then SuggDY = SuggDY + 1
                If .SolidTest(CurX + W - 1, CurY + H) Then SuggDX = SuggDX - 1.5
                If .SolidTest(CurX, CurY + H) Then SuggDX = SuggDX - 1
                If .SolidTest(CurX - 1, CurY + H - 1) Then SuggDY = SuggDY - 1.5
                If .SolidTest(CurX - 1, CurY) Then SuggDY = SuggDY - 1
                If SuggDX = 0 And SuggDY = 0 Then
                    If .SolidTest(CurX - 1, CurY + H) Then SuggDX = SuggDX - 1
                    If .SolidTest(CurX - 1, CurY - 1) Then SuggDY = SuggDY - 1
                    If .SolidTest(CurX + W, CurY - 1) Then SuggDX = SuggDX + 1
                    If .SolidTest(CurX + W, CurY + H) Then SuggDY = SuggDY + 1
                End If
            End With
            Exit Sub
        End If
    End With

    SuggDX = Sgn(CurDX)
    SuggDY = Sgn(CurDY)
    If BSol Then
        If Not LSol Then
            RotateMotion SuggDX, SuggDY, False
        ElseIf OSol Then
            RotateMotion SuggDX, SuggDY, True
        End If
        bLost = False
    Else
        If OSol Then
            RotateMotion SuggDX, SuggDY, True
            bLost = False
        ElseIf TSol Then
            bLost = False
        Else
            bLost = True
        End If
    End If
End Sub

Private Sub GetSuggestedRightWallDeltas(ByVal CurX As Single, ByVal CurY As Single, ByVal CurDX As Single, ByVal CurDY As Single, SuggDX As Single, SuggDY As Single, bLost As Boolean)
    Dim W As Integer, H As Integer
    Dim BSol As Single, TSol As Single, OSol As Single
    Dim LSol As Single
    
    SuggDX = 0
    SuggDY = 0
    W = Width
    H = Height

    With rDef
        If CurDX = 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX + W, CurY + H)
            LSol = .SolidTest(CurX + W, CurY + H - 2)
            TSol = .SolidTest(CurX + W, CurY)
            OSol = .SolidTest(CurX + W - 1, CurY - 1)
        ElseIf CurDX > 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX + W - 1, CurY + H)
            LSol = .SolidTest(CurX + W, CurY + H - 1)
            TSol = LSol
            OSol = .SolidTest(CurX + W, CurY)
        ElseIf CurDX > 0 And CurDY = 0 Then
            BSol = .SolidTest(CurX - 1, CurY + H)
            LSol = .SolidTest(CurX + 1, CurY + H)
            TSol = .SolidTest(CurX + W - 1, CurY + H)
            OSol = .SolidTest(CurX + W, CurY + H - 1)
        ElseIf CurDX > 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX - 1, CurY + H - 1)
            LSol = .SolidTest(CurX, CurY + H)
            TSol = LSol
            OSol = .SolidTest(CurX + W - 1, CurY + H)
        ElseIf CurDX = 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX - 1, CurY - 1)
            LSol = .SolidTest(CurX - 1, CurY + 1)
            TSol = .SolidTest(CurX - 1, CurY + H - 1)
            OSol = .SolidTest(CurX, CurY + H)
        ElseIf CurDX < 0 And CurDY > 0 Then
            BSol = .SolidTest(CurX, CurY - 1)
            LSol = .SolidTest(CurX - 1, CurY)
            TSol = LSol
            OSol = .SolidTest(CurX - 1, CurY + H - 1)
        ElseIf CurDX < 0 And CurDY = 0 Then
            BSol = .SolidTest(CurX + W, CurY - 1)
            LSol = .SolidTest(CurX + W - 2, CurY - 1)
            TSol = .SolidTest(CurX, CurY - 1)
            OSol = .SolidTest(CurX - 1, CurY)
        ElseIf CurDX < 0 And CurDY < 0 Then
            BSol = .SolidTest(CurX + W, CurY)
            LSol = .SolidTest(CurX + W - 1, CurY - 1)
            TSol = LSol
            OSol = .SolidTest(CurX, CurY - 1)
        Else ' cury=0 and curx=0
            With rDef
                If .SolidTest(CurX, CurY - 1) Then SuggDX = SuggDX - 1
                If .SolidTest(CurX + W - 1, CurY - 1) Then SuggDX = SuggDX - 1.5
                If .SolidTest(CurX + W, CurY) Then SuggDY = SuggDY - 1
                If .SolidTest(CurX + W, CurY + H - 1) Then SuggDY = SuggDY - 1.5
                If .SolidTest(CurX + W - 1, CurY + H) Then SuggDX = SuggDX + 1
                If .SolidTest(CurX, CurY + H) Then SuggDX = SuggDX + 1.5
                If .SolidTest(CurX - 1, CurY + H - 1) Then SuggDY = SuggDY + 1
                If .SolidTest(CurX - 1, CurY) Then SuggDY = SuggDY + 1.5
                If SuggDX = 0 And SuggDY = 0 Then
                    If .SolidTest(CurX - 1, CurY + H) Then SuggDY = SuggDY + 1
                    If .SolidTest(CurX - 1, CurY - 1) Then SuggDX = SuggDX - 1
                    If .SolidTest(CurX + W, CurY - 1) Then SuggDY = SuggDY - 1
                    If .SolidTest(CurX + W, CurY + H) Then SuggDX = SuggDX + 1
                End If
            End With
            Exit Sub
        End If
    End With

    SuggDX = Sgn(CurDX)
    SuggDY = Sgn(CurDY)
    If BSol Then
        If Not LSol Then
            RotateMotion SuggDX, SuggDY, True
        ElseIf OSol Then
            RotateMotion SuggDX, SuggDY, False
        End If
        bLost = False
    Else
        If OSol Then
            RotateMotion SuggDX, SuggDY, False
            bLost = False
        ElseIf TSol Then
            bLost = False
        Else
            bLost = True
        End If
    End If
End Sub

Private Sub RotateMotion(RDX As Single, RDY As Single, ByVal bRight As Boolean)
    Dim Tmp As Single
    
    If bRight Then
        Tmp = Sgn(RDX + RDY)
        RDX = Sgn(RDX - RDY)
        RDY = Tmp
    Else
        Tmp = Sgn(RDY - RDX)
        RDX = Sgn(RDX + RDY)
        RDY = Tmp
    End If
End Sub
