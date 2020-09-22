VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpriteDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines a sprite based on a path and a template.  Indexed and owned by map. (Path determines layer.)"
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: SpritDef.cls - Sprite Definition Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name of this SpriteDef as indexed by the Map."
Public rLayer As Layer
Attribute rLayer.VB_VarProcData = ";Relational"
Attribute rLayer.VB_VarDescription = "Reference to the layer where sprites of this SpriteDef exist (comes from rPath)."
Public Flags As Byte
Attribute Flags.VB_VarDescription = "eDefFlagBits value (create instance on map start?)"
Public rPath As Path
Attribute rPath.VB_VarProcData = ";Relational"
Attribute rPath.VB_VarDescription = "Reference to the path that sprites of this SpriteDef follow."
Public Template As SpriteTemplate
Attribute Template.VB_VarProcData = ";Relational"
Attribute Template.VB_VarDescription = "Reference to the SpriteTemplate object used by this  SpriteDef."

Public Enum eDefFlagBits
    FLAG_INSTANCE = 1
End Enum

Public Function Clone() As SpriteDef
Attribute Clone.VB_Description = "Create and return a copy of this SpriteDef."
    Dim C As New SpriteDef
    
    C.Name = Name
    Set C.rLayer = rLayer
    C.Flags = Flags
    Set C.rPath = rPath
    Set C.Template = Template
    Set Clone = C
End Function

Public Sub LoadTiles(Disp As BMDXDisplay)
Attribute LoadTiles.VB_Description = "Load into display memory all the tilesets required to draw sprites of this SpriteDef."
    Template.LoadTiles Disp
End Sub

Property Get StateCount() As Integer
Attribute StateCount.VB_Description = "Return the number of sprite states in this SpriteDef's template."
    StateCount = Template.StateCount
End Property

Public Property Get StateFrame(ByVal State As Integer, ByVal Frame As Integer) As Integer
Attribute StateFrame.VB_Description = "Return the tile index for the specified frame of the specified state of the SpriteTemplate."
    StateFrame = Template.StateFrame(State, Frame)
End Property

Public Property Get StateFrameCount(ByVal State As Integer) As Integer
Attribute StateFrameCount.VB_Description = "Return the number of frames for the specified state in the SpriteDef's Template."
    StateFrameCount = Template.StateFrameCount(State)
End Property

Public Sub ReleaseRefs()
Attribute ReleaseRefs.VB_Description = "Release all references this SpriteDef has to other objects (to avoid circular reference problems)."
    
    Set rLayer = Nothing
    Set rPath = Nothing
    Set Template = Nothing
    
End Sub

Public Function MakeInstance() As Sprite
Attribute MakeInstance.VB_Description = "Create and return a new instance of this sprite (must be added to layer to be automatically processed)."
    Dim S As New Sprite
    
    If Not (rPath Is Nothing) Then
        S.X = rPath.PointX(0)
        S.Y = rPath.PointY(0)
    End If
    Set S.rDef = Me
    Set MakeInstance = S
    S.CurFrame = 0
    
End Function

Public Function SolidTest(ByVal XPos As Long, ByVal YPos As Long) As Boolean
Attribute SolidTest.VB_Description = "Test whether the layer is solid at the given coordinate, according to the template's solidity definition."
    Dim Tile As Integer
    Dim W As Integer, H As Integer
    Dim WR As Integer, HR As Integer
    
    If Template.SolidInfo Is Nothing Then Exit Function
    
    W = rLayer.TSDef.TileWidth
    H = rLayer.TSDef.TileHeight
    If XPos < 0 Or YPos < 0 Or (Int(XPos / W) >= rLayer.Columns) Or (Int(YPos / H) >= rLayer.Rows) Then
        SolidTest = True
        Exit Function
    End If
    
    Tile = rLayer.Data.TileValue(Int(XPos / W), Int(YPos / H))
    
    With Template.SolidInfo
        If Not (.Solid Is Nothing) Then
            If .Solid.Group.IsMember(Tile) Then
                SolidTest = True
                Exit Function
            End If
        End If
    
        WR = XPos Mod W
        HR = YPos Mod H
        
        If Not (.Uphill Is Nothing) Then
            If .Uphill.Group.IsMember(Tile) Then
                If WR + HR + 1 >= H Then
                    SolidTest = True
                End If
                Exit Function
             End If
        End If
        
        If Not (.Downhill Is Nothing) Then
            If .Downhill.Group.IsMember(Tile) Then
                If WR <= HR Then
                    SolidTest = True
                End If
                Exit Function
            End If
        End If
        
        If Not (.UpCeil Is Nothing) Then
            If .UpCeil.Group.IsMember(Tile) Then
                If WR + HR < H Then
                    SolidTest = True
                End If
                Exit Function
            End If
        End If
        
        If Not (.DownCeil Is Nothing) Then
            If .DownCeil.Group.IsMember(Tile) Then
                If WR >= HR Then
                    SolidTest = True
                End If
                Exit Function
            End If
        End If
    End With
    
End Function

Public Property Get StateTilesetDef(ByVal State As Integer) As TileSetDef
Attribute StateTilesetDef.VB_Description = "Return the TileSetDef object used by the specified sprite state in the SpriteDef's SpriteTemplate."
    Set StateTilesetDef = Template.StateTilesetDef(State)
End Property

Public Function Save(ByVal nFileNum As Integer)
Attribute Save.VB_Description = "Save this SpriteDef to a binary (map) file."
    Dim I As Integer
    Dim J As Integer
    Dim bFound As Boolean
    
    SaveString nFileNum, Name
    bFound = False
    For I = 0 To Prj.MapCount - 1
        For J = 0 To Prj.Maps(I).LayerCount - 1
            If rLayer Is Prj.Maps(I).MapLayer(J) Then
                Put #nFileNum, , I
                Put #nFileNum, , J
                bFound = True
            End If
        Next
    Next
    If Not bFound Then Err.Raise vbObjectError, , "Cannot find layer " & rLayer.Name & " for sprite definition " & Name
    Put #nFileNum, , Flags
    SaveString nFileNum, rPath.Name
    If Template Is Nothing Then Put #nFileNum, , -1 Else Put #nFileNum, , Template.GetIndexes(I)
    
End Function

Public Function Load(ByVal nFileNum As Integer)
Attribute Load.VB_Description = "Load a SpriteDef from a binary (map) file."
    Dim I As Integer
    Dim J As Integer
    Dim MapIdx As Integer

    Name = LoadString(nFileNum)
    
    Get #nFileNum, , MapIdx
    Get #nFileNum, , J
    Set rLayer = Prj.Maps(MapIdx).MapLayer(J)
    Get #nFileNum, , Flags
    Set rPath = Prj.Maps(MapIdx).Paths(LoadString(nFileNum))
    Get #nFileNum, , I
    If I >= 0 Then Set Template = Prj.Maps(MapIdx).SpriteTemplates(I) Else Set Template = Nothing
    
End Function

