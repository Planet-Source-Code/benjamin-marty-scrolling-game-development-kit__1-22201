VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpriteTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines detailed parameters about a type of sprite including motion, graphics, states, collisions, etc. Indexed and owned by map."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: SprTmplt.cls - Sprite Template Definition Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name of this SpriteTemplate as indexed by the map."
Public SolidInfo As SolidDef
Attribute SolidInfo.VB_VarProcData = ";Relational"
Attribute SolidInfo.VB_VarDescription = "Reference to the SolidDef describing which tile indexes on a sprite's layer it sees as solid."
Public AnimSpeed As Byte ' 0=fastest
Attribute AnimSpeed.VB_VarProcData = ";Behavior"
Attribute AnimSpeed.VB_VarDescription = "How many frames of a particular image are displayed before switching to the next image."
Private StateDefs() As State
Public StateType As eStateType
Attribute StateType.VB_VarProcData = ";Behavior"
Attribute StateType.VB_VarDescription = "Indicates the type/number of states this sprite contains from a pre-defined set supported by Sprite.Advance."
Public ControlType As eControlType
Attribute ControlType.VB_VarProcData = ";Behavior"
Attribute ControlType.VB_VarDescription = "Indicates what controls the motion of this sprite."
Public MoveSpeed As Byte
Attribute MoveSpeed.VB_VarProcData = ";Behavior"
Attribute MoveSpeed.VB_VarDescription = "Maximum velocity of the sprite in pixels per frame."
Public GravPow As Byte
Attribute GravPow.VB_VarProcData = ";Behavior"
Attribute GravPow.VB_VarDescription = "Gravity: 0=strong up, 10=none, 20=strong down."
Public Inertia As Byte
Attribute Inertia.VB_VarProcData = ";Behavior"
Attribute Inertia.VB_VarDescription = "Percent of DX, DY (velocity) retained by a sprite after each frame (0-100) when not accelerating."
Public Flags As Integer
Attribute Flags.VB_VarProcData = ";Behavior"
Attribute Flags.VB_VarDescription = "Combined bits from eTemplateFlagBits."
Public CollClass As Integer
Attribute CollClass.VB_VarProcData = ";Behavior"
Attribute CollClass.VB_VarDescription = "Each bit (n=0-15) represents whether sprites of this template are a member of collision class #n."

Private Type State
    Frames As Variant
    TSDef As TileSetDef
End Type

Public Enum eStateType
    STATE_SINGLE = 1
    STATE_LEFT_RIGHT
    STATE_8_DIRECTION
    STATE_36_DIRECTION
End Enum

Public Enum eTemplateFlagBits
    FLAG_ACCELSTATES = 1
    FLAG_UPNEEDSSOLID = 4
End Enum

Public Enum eControlType
   CONTROL_APPROX_PATH = 1
   CONTROL_EXACT_PATH
   CONTROL_INPUT
   CONTROL_FOLLOWNEAR
   CONTROL_FOLLOWPLAYER
   CONTROL_FOLLOWFLATFLOOR
   CONTROL_FOLLOWSLOPEFLOOR
   CONTROL_FOLLOWFLATWALL
   CONTROL_FOLLOWLEFTWALL
   CONTROL_FOLLOWRIGHTWALL
   CONTROL_SLOPEFLOORJUMP
   CONTROL_INERT
   CONTROL_SIMPLE
End Enum

Public Function Clone() As SpriteTemplate
Attribute Clone.VB_Description = "Return a copy of this SpriteTemplate."
    Dim C As New SpriteTemplate
    Dim I As Integer
    Dim J As Integer
    
    C.Name = Name
    C.Flags = Flags
    Set C.SolidInfo = SolidInfo
    C.AnimSpeed = AnimSpeed
    C.StateCount = StateCount
    For I = 0 To StateCount - 1
       Set C.StateTilesetDef(I) = StateTilesetDef(I)
       For J = 0 To StateFrameCount(I) - 1
          C.AppendStateFrame I, Me.StateFrame(I, J)
       Next
    Next
    C.StateType = StateType
    C.ControlType = ControlType
    C.MoveSpeed = MoveSpeed
    C.GravPow = GravPow
    C.Inertia = Inertia
    C.CollClass = CollClass
    
    Set Clone = C
End Function

Public Sub LoadTiles(Disp As BMDXDisplay)
Attribute LoadTiles.VB_Description = "Load into display memory all the graphics required to display sprites of this template."
    Dim I As Integer
    
    If StateCount = 0 Then Exit Sub
    
    For I = LBound(StateDefs) To UBound(StateDefs)
        If Not (StateDef(I).TSDef Is Nothing) Then
            If Not StateDef(I).TSDef.IsLoaded Then
                StateDefs(I).TSDef.Load
            End If
            If Not StateDef(I).TSDef.IsLoaded Then
                Err.Raise vbObjectError, , "Failed to load " & StateDef(I).TSDef.Name & " (" & StateDef(I).TSDef.ImagePath & ")"
            End If
        
            If StateDef(I).TSDef.LoadedInstance Is Nothing Then
                Set StateDefs(I).TSDef.LoadedInstance = Disp.CreateTileSet(StateDefs(I).TSDef.Image, StateDefs(I).TSDef.TileWidth, StateDefs(I).TSDef.TileHeight)
            End If
        End If
    Next I
    
End Sub

Property Let StateCount(ByVal Cnt As Integer)
Attribute StateCount.VB_Description = "Set/get the number of sprite states in this template."
    If Cnt > 0 Then
        ReDim Preserve StateDefs(0 To Cnt - 1) As State
    Else
        Erase StateDefs
    End If
End Property

Property Get StateCount() As Integer
    On Error GoTo CountZero
    
    StateCount = UBound(StateDefs) - LBound(StateDefs) + 1
    Exit Property
    
CountZero:
    StateCount = 0
End Property

Private Property Get StateDef(ByVal Index As Integer) As State
    StateDef = StateDefs(Index)
End Property

Property Set StateTilesetDef(ByVal State As Integer, TSDef As TileSetDef)
Attribute StateTilesetDef.VB_Description = "Set/get the TileSetDef containing graphics for the specified state's frames."
Attribute StateTilesetDef.VB_ProcData.VB_Invoke_PropertyPutRef = ";Relational"
    Set StateDefs(State).TSDef = TSDef
End Property

Public Sub AppendStateFrame(ByVal State As Integer, ByVal Frame As Integer)
Attribute AppendStateFrame.VB_Description = "Add a new tile to the animation/image for a particular sprite state."
    If IsEmpty(StateDef(State).Frames) Then
        ReDim StateDefs(State).Frames(0) As Integer
    Else
        ReDim Preserve StateDefs(State).Frames(LBound(StateDef(State).Frames) To UBound(StateDef(State).Frames) + 1)
    End If
    
    StateDefs(State).Frames(UBound(StateDef(State).Frames)) = Frame
End Sub

Public Property Get StateFrame(ByVal State As Integer, ByVal Frame As Integer) As Integer
Attribute StateFrame.VB_Description = "Return the tile index for the specified frame of the specified state."
Attribute StateFrame.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If IsEmpty(StateDef(State).Frames) Then
        StateFrame = -1
    Else
        StateFrame = StateDef(State).Frames(Frame)
    End If
End Property

Public Property Get StateFrameCount(ByVal State As Integer) As Integer
Attribute StateFrameCount.VB_Description = "Return the number of animation frames for the specified sprite state."
   If IsEmpty(StateDef(State).Frames) Then
      StateFrameCount = 0
   Else
      StateFrameCount = UBound(StateDef(State).Frames) - LBound(StateDef(State).Frames) + 1
   End If
End Property

Public Sub ClearState(ByVal State As Integer)
Attribute ClearState.VB_Description = "Erase all frames from a particular sprite state animation/image."
    StateDefs(State).Frames = Empty
End Sub

Public Sub ReleaseRefs()
Attribute ReleaseRefs.VB_Description = "Release all references this SpriteTemplate has to other objects (to avoid circular reference problems)."
    Dim I As Integer
    
    For I = 0 To StateCount - 1
        Set StateDefs(I).TSDef = Nothing
    Next I
    
    Set SolidInfo = Nothing
End Sub

Public Property Get StateTilesetDef(ByVal State As Integer) As TileSetDef
    Set StateTilesetDef = StateDef(State).TSDef
End Property

Public Sub Save(ByVal nFileNum As Integer)
Attribute Save.VB_Description = "Save this SpriteTemplate to a binary (map) file."
    Dim I As Integer
    Dim J As Integer
    Dim TmpInt As Integer
    Dim TmpBy As Byte
    
    SaveString nFileNum, Name
    If Not (SolidInfo Is Nothing) Then
        SaveString nFileNum, SolidInfo.Name
        SaveString nFileNum, SolidInfo.TSName
    Else
        SaveString nFileNum, ""
        SaveString nFileNum, ""
    End If

    Put #nFileNum, , Flags
    Put #nFileNum, , AnimSpeed
    Put #nFileNum, , MoveSpeed
    Put #nFileNum, , GravPow
    Put #nFileNum, , Inertia
    Put #nFileNum, , CollClass
    TmpInt = StateCount
    Put #nFileNum, , TmpInt
    TmpBy = StateType
    Put #nFileNum, , TmpBy
    TmpBy = ControlType
    Put #nFileNum, , TmpBy

    For I = 0 To StateCount - 1
        If StateDefs(I).TSDef Is Nothing Then
            TmpInt = -1
        Else
            TmpInt = Prj.TileSetDef(StateDefs(I).TSDef.Name).GetIndex
        End If
        Put #nFileNum, , TmpInt
        TmpInt = StateFrameCount(I)
        Put #nFileNum, , TmpInt
        For J = 0 To StateFrameCount(I) - 1
            TmpInt = StateFrame(I, J)
            Put #nFileNum, , TmpInt
        Next J
    Next I

End Sub

Public Sub Load(ByVal nFileNum As Integer)
Attribute Load.VB_Description = "Load a SpriteTemplate from a binary (map) file."
    Dim I As Integer
    Dim J As Integer
    Dim nFrCnt As Integer
    Dim F As Integer
    Dim SolidInfoName As String
    Dim SolidInfoTileset As String
    Dim nStateCount As Integer
    Dim TmpBy As Byte
    Dim TmpInt As Integer

    Name = LoadString(nFileNum)
    SolidInfoName = LoadString(nFileNum)
    SolidInfoTileset = LoadString(nFileNum)
    
    Get #nFileNum, , Flags
    Get #nFileNum, , AnimSpeed
    Get #nFileNum, , MoveSpeed
    Get #nFileNum, , GravPow
    Get #nFileNum, , Inertia
    Get #nFileNum, , CollClass
    Get #nFileNum, , nStateCount
    Get #nFileNum, , TmpBy
    StateType = TmpBy
    Get #nFileNum, , TmpBy
    ControlType = TmpBy
    
    If (Len(SolidInfoName) > 0) And (Len(SolidInfoTileset) > 0) Then
        Set SolidInfo = Prj.SolidDefs(SolidInfoName, SolidInfoTileset)
    Else
        Set SolidInfo = Nothing
    End If
    
    StateCount = nStateCount
    For I = 0 To StateCount - 1
        Get #nFileNum, , TmpInt
        If TmpInt >= 0 Then Set StateTilesetDef(I) = Prj.TileSetDef(TmpInt)
        Get #nFileNum, , nFrCnt
        For J = 0 To nFrCnt - 1
            Get #nFileNum, , F
            AppendStateFrame I, F
        Next
    Next
End Sub

Public Function GetIndexes(ByRef IndexOfMap As Integer) As Integer
Attribute GetIndexes.VB_Description = "Return the index of this template as indexed by the map, and set IndexOfMap to the index of the sprite's map within the project."
    Dim I As Integer
    
    For IndexOfMap = 0 To Prj.MapCount - 1
        For I = 0 To Prj.Maps(IndexOfMap).SpriteTemplateCount - 1
            If Prj.Maps(IndexOfMap).SpriteTemplates(I) Is Me Then
                GetIndexes = I
                Exit Function
            End If
        Next I
    Next IndexOfMap
    GetIndexes = -1
    IndexOfMap = -1
End Function

Private Sub Class_Initialize()
    GravPow = 10
End Sub
