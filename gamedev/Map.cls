VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Map, corresponds mainly to one MAP file. Indexed and owned by project."
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
' File: Map.cls - Map Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this Map object as indexed by the project."
Dim DataFilePath As String
Dim Layers As Variant
Dim PixWidth As Long
Dim PixHeight As Long
Dim DrawLeft As Integer
Dim DrawTop As Integer
Dim DrawWidth As Integer
Dim DrawHeight As Integer
Dim PathColl As Variant
Dim SpriteDefColl As Variant
Dim SprTplColl As Variant
Dim CollDefColl As Variant
Dim CollClassNames(0 To 15) As String
Dim SpecialColl As Variant
Dim InteractColl As Variant

Public WithEvents Disp As BMDXDisplay
Attribute Disp.VB_VarHelpID = -1
Public bActive As Boolean
Attribute bActive.VB_VarDescription = "Determines whether the map is updating.  (Auto set to False when display is minimized.)"
Public IsDirty As Boolean
Attribute IsDirty.VB_VarDescription = "Returns true if map data has changed."
Public PlayerSpriteName As String
Public BackgroundMusic As String
Attribute BackgroundMusic.VB_VarProcData = ";Relational"
Attribute BackgroundMusic.VB_VarDescription = "Name of the MediaClip containing the background music for this map."
Public bDisablePlayerEdit As Boolean

Public Sub ReleaseRefs()
Attribute ReleaseRefs.VB_Description = "Release references this map has to other objects to avoid circular reference problems."
    Layers = Empty
    PathColl = Empty
    SpriteDefColl = Empty
    SprTplColl = Empty
    CollDefColl = Empty
    SpecialColl = Empty
    Set Disp = Nothing
End Sub

Public Sub ReleaseSpriteRefs()
Attribute ReleaseSpriteRefs.VB_Description = "Release references all SpriteDefs in the map have to other objects to avoid circular reference problems."
    Dim Idx As Integer
    
    For Idx = 0 To SpriteDefCount - 1
        SpriteDefs(Idx).ReleaseRefs
    Next
End Sub

Public Sub ReleaseSpriteTemplateRefs()
Attribute ReleaseSpriteTemplateRefs.VB_Description = "Release references all SpriteTemplates in the map have to other objects to avoid circular reference problems."
    Dim Idx As Integer
    
    For Idx = 0 To SpriteTemplateCount - 1
        SpriteTemplates(Idx).ReleaseRefs
    Next
End Sub

Public Sub ReleaseLayerRefs()
Attribute ReleaseLayerRefs.VB_Description = "Release the references all layers have to other objects to avoid circular reference problems."
    Dim Idx As Integer
    
    For Idx = 0 To LayerCount - 1
        MapLayer(Idx).ReleaseRefs
    Next
End Sub

Sub AddLayer(ByVal Name As String, ByVal TS As String, ByVal XR As Single, ByVal YR As Single, ByVal Trans As Boolean)
Attribute AddLayer.VB_Description = "Add a new Layer to the map with the specified parameters."
    If IsEmpty(Layers) Then
        ReDim Layers(0) As Layer
    Else
        ReDim Preserve Layers(LBound(Layers) To UBound(Layers) + 1) As Layer
    End If
    
    Set Layers(UBound(Layers)) = New Layer
    With Layers(UBound(Layers))
        .Create Name, PixWidth, PixHeight, Prj.TileSetDef(TS), XR, YR, Trans, Me
    End With
    
    IsDirty = True
End Sub

Sub RemoveLayer(ByVal Key As Variant)
Attribute RemoveLayer.VB_Description = "Remove a layer and all related Paths, SpriteDefs and SpecialFunctions from the map."
    Dim Index As Integer
    Dim RmvIdx As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(Layers) To UBound(Layers)
            If Layers(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(Layers) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveLayer"
    End If
    
    If PathCount > 0 Then
        RmvIdx = 0
        Do
            If Paths(RmvIdx).LayerName = Layers(Index).Name Then
                RemovePath RmvIdx
            Else
                RmvIdx = RmvIdx + 1
            End If
        Loop Until RmvIdx >= PathCount
    End If
    
    If SpecialCount > 0 Then
        RmvIdx = 0
        Do
            If Specials(RmvIdx).LayerIndex = Index Then
                RemoveSpecial RmvIdx
            Else
                If Specials(RmvIdx).LayerIndex > Index Then
                    Specials(RmvIdx).LayerIndex = Specials(RmvIdx).LayerIndex - 1
                End If
                RmvIdx = RmvIdx + 1
            End If
        Loop Until RmvIdx >= SpecialCount
    End If
    
    Do While Index < UBound(Layers)
        Set Layers(Index) = Layers(Index + 1)
        Index = Index + 1
    Loop
    Set Layers(UBound(Layers)) = Nothing
    
    If Index > LBound(Layers) Then
        ReDim Preserve Layers(LBound(Layers) To UBound(Layers) - 1)
    Else
        Layers = Empty
    End If

    IsDirty = True
End Sub

Public Sub AddPath(P As Path)
Attribute AddPath.VB_Description = "Add a new Path to the map."
    If IsEmpty(PathColl) Then
        ReDim PathColl(0) As Path
    Else
        ReDim Preserve PathColl(LBound(PathColl) To UBound(PathColl) + 1)
    End If
    Set PathColl(UBound(PathColl)) = P
    IsDirty = True
End Sub

Public Sub RemovePath(ByVal Key As Variant)
Attribute RemovePath.VB_Description = "Remove a path and all related SpriteDefs from the map."
    Dim Index As Integer
    Dim RmvIdx As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(PathColl) To UBound(PathColl)
            If PathColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(PathColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemovePath"
    End If
    
    RmvIdx = 0
    If SpriteDefCount > 0 Then
        Do
            If SpriteDefs(RmvIdx).rPath Is PathColl(Index) Then
                RemoveSpriteDef RmvIdx
            Else
                RmvIdx = RmvIdx + 1
            End If
        Loop Until RmvIdx >= SpriteDefCount
    End If
    
    Do While Index < UBound(PathColl)
        Set PathColl(Index) = PathColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set PathColl(UBound(PathColl)) = Nothing
    
    If Index > LBound(PathColl) Then
        ReDim Preserve PathColl(LBound(PathColl) To UBound(PathColl) - 1)
    Else
        PathColl = Empty
    End If
    IsDirty = True
End Sub

Property Get Paths(ByVal Key As Variant) As Path
Attribute Paths.VB_Description = "Return a Path object given its name or ordinal index in the map."
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(PathColl) To UBound(PathColl)
            If UCase$(PathColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(PathColl) Then
        Err.Raise vbObjectError, , "Path " & CStr(Key) & " not found"
    End If
    
    Set Paths = PathColl(I)
End Property

Public Function PathExists(ByVal Key As Variant) As Boolean
Attribute PathExists.VB_Description = "Return True if the specified path name or index exists in this map."
    Dim I As Integer
    
    If IsEmpty(PathColl) Then
        PathExists = False
        Exit Function
    End If
    
    If VarType(Key) = vbString Then
        For I = LBound(PathColl) To UBound(PathColl)
            If UCase$(PathColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(PathColl) Then
        PathExists = False
    Else
        PathExists = True
    End If
    
End Function

Public Property Get PathCount() As Integer
Attribute PathCount.VB_Description = "Return the total number of Path objects in this map."
    If IsEmpty(PathColl) Then
        PathCount = 0
    Else
        PathCount = UBound(PathColl) - LBound(PathColl) + 1
    End If
End Property

Public Sub AddSpecial(SpecFunc As SpecialFunction)
Attribute AddSpecial.VB_Description = "Add a new SpecialFunction to the map."
    If IsEmpty(SpecialColl) Then
        ReDim SpecialColl(0) As SpecialFunction
    Else
        ReDim Preserve SpecialColl(LBound(SpecialColl) To UBound(SpecialColl) + 1)
    End If
    Set SpecialColl(UBound(SpecialColl)) = SpecFunc
    IsDirty = True
End Sub

Public Sub RemoveSpecial(ByVal Key As Variant)
Attribute RemoveSpecial.VB_Description = "Remove a SpecialFunction from the map."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(SpecialColl) To UBound(SpecialColl)
            If SpecialColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(SpecialColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveSpecial"
    End If
    
    Do While Index < UBound(SpecialColl)
        Set SpecialColl(Index) = SpecialColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set SpecialColl(UBound(SpecialColl)) = Nothing
    
    If Index > LBound(SpecialColl) Then
        ReDim Preserve SpecialColl(LBound(SpecialColl) To UBound(SpecialColl) - 1)
    Else
        SpecialColl = Empty
    End If
    IsDirty = True
End Sub

Property Get Specials(ByVal Key As Variant) As SpecialFunction
Attribute Specials.VB_Description = "Return a SpecialFunction given its name or ordinal index in the Map."
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(SpecialColl) To UBound(SpecialColl)
            If UCase$(SpecialColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(SpecialColl) Then
        Err.Raise vbObjectError, , "Special function " & CStr(Key) & " not found"
    End If
    
    Set Specials = SpecialColl(I)
End Property

Public Function SpecialExists(ByVal Key As Variant) As Boolean
Attribute SpecialExists.VB_Description = "Return True if the SpecialFunction name or index exists in the map."
    Dim I As Integer
    
    If IsEmpty(SpecialColl) Then
        SpecialExists = False
        Exit Function
    End If
    
    If VarType(Key) = vbString Then
        For I = LBound(SpecialColl) To UBound(SpecialColl)
            If UCase$(SpecialColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(SpecialColl) Then
        SpecialExists = False
    Else
        SpecialExists = True
    End If
    
End Function

Public Property Get SpecialCount() As Integer
Attribute SpecialCount.VB_Description = "Return the number of SpecialFunctions in the map."
    If IsEmpty(SpecialColl) Then
        SpecialCount = 0
    Else
        SpecialCount = UBound(SpecialColl) - LBound(SpecialColl) + 1
    End If
End Property

Public Sub AddCollDef(CollDef As CollisionDef)
Attribute AddCollDef.VB_Description = "Add a new CollDef object to the map."
    If IsEmpty(CollDefColl) Then
        ReDim CollDefColl(0) As CollisionDef
    Else
        ReDim Preserve CollDefColl(LBound(CollDefColl) To UBound(CollDefColl) + 1) As CollisionDef
    End If
    Set CollDefColl(UBound(CollDefColl)) = CollDef
    IsDirty = True
End Sub

Public Sub RemoveCollDef(ByVal Index As Integer)
Attribute RemoveCollDef.VB_Description = "Remove a CollisionDef object from the map."
    Dim I As Integer
    If Not IsEmpty(CollDefColl) Then
        If Index >= LBound(CollDefColl) And Index <= UBound(CollDefColl) Then
            For I = Index To UBound(CollDefColl) - 1
                Set CollDefColl(I) = CollDefColl(I + 1)
            Next
            Set CollDefColl(UBound(CollDefColl)) = Nothing
            If UBound(CollDefColl) > 0 Then
                ReDim Preserve CollDefColl(LBound(CollDefColl) To UBound(CollDefColl) - 1)
            Else
                CollDefColl = Empty
            End If
        End If
        IsDirty = True
    End If
End Sub

Public Property Get CollDefCount() As Integer
Attribute CollDefCount.VB_Description = "Return the number of collision definitions for sprites on the map."
    If IsEmpty(CollDefColl) Then
        CollDefCount = 0
    Else
        CollDefCount = UBound(CollDefColl) - LBound(CollDefColl) + 1
    End If
End Property

Public Property Get CollDefs(ByVal Index As Integer) As CollisionDef
Attribute CollDefs.VB_Description = "Return a CollisionDef object given its ordinal index in the Map."
    If Not IsEmpty(CollDefColl) Then
        If Index >= LBound(CollDefColl) And Index <= UBound(CollDefColl) Then
            Set CollDefs = CollDefColl(Index)
        End If
    End If
End Property

Public Property Let CollClassName(ByVal Index As Integer, ByVal NewVal As String)
Attribute CollClassName.VB_Description = "Get or set the name of one of the 16 collision classes for sprites."
Attribute CollClassName.VB_ProcData.VB_Invoke_PropertyPut = ";Text"
    CollClassNames(Index) = NewVal
End Property

Public Property Get CollClassName(ByVal Index As Integer) As String
    CollClassName = CollClassNames(Index)
    If CollClassName = "" Then CollClassName = "Class " & CStr(Index + 1)
End Property

Sub UnloadTiles()
Attribute UnloadTiles.VB_Description = "Remove from memory and display memory all graphics that are saved."
    Dim I As Integer
    Dim J As Integer
    
    On Error Resume Next
    
    If (Not (Disp Is Nothing)) And (Not IsEmpty(Layers)) Then
        For I = 0 To Prj.TileSetDefCount - 1
            J = 0
            Do While J < Disp.TileSets.Count
                If Prj.TileSetDef(I).LoadedInstance Is Disp.TileSets.Item(J) Then
                    Disp.TileSets.Remove J
                Else
                    J = J + 1
                End If
            Loop
            If Len(Prj.TileSetDef(I).ImagePath) > 0 And Prj.TileSetDef(I).IsDirty = False Then Prj.TileSetDef(I).Unload
            Set Prj.TileSetDef(I).LoadedInstance = Nothing
        Next I
        For I = LBound(Layers) To UBound(Layers)
            Layers(I).UnloadTiles
        Next
    End If

End Sub

Private Sub Disp_Close()
    UnloadTiles
    Set Disp = Nothing
    Set CurDisp = Nothing
End Sub

Private Sub Disp_TileSetLost(ByVal Index As Integer)
    Dim I As Integer
    On Error GoTo RestoreErr
    
    If IsEmpty(Layers) Then Exit Sub
    For I = 0 To Prj.TileSetDefCount - 1
        If Disp.TileSets(Index) Is Prj.TileSetDef(I).LoadedInstance Then
            bActive = True
            With Prj.TileSetDef(I)
                .LoadedInstance.Restore
                If Not .IsLoaded Then
                    .Load
                End If
                .LoadedInstance.PaintPicture .Image, 0, 0
            End With
        End If
    Next I
    Exit Sub
    
RestoreErr:
    bActive = False
End Sub

Sub LoadTiles()
Attribute LoadTiles.VB_Description = "Load the tilesets for all sprites and layers of this map into display memory."
    Dim I As Integer
    
    If Disp Is Nothing Then
        Err.Raise vbObjectError, , "Cannot load map layer graphics without a display"
    End If
    
    If IsEmpty(Layers) Then Exit Sub
    
    For I = LBound(Layers) To UBound(Layers)
        Layers(I).LoadTiles Disp
    Next I
    
    If Not IsEmpty(SpriteDefColl) Then
        For I = LBound(SpriteDefColl) To UBound(SpriteDefColl)
            SpriteDefs(I).LoadTiles Disp
        Next I
    End If
    
End Sub

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the map's parameters to a plain text string for the GDP project file."
    
    Serialize = GeneralSerialize("Name", "DataFilePath", "DrawLeft", "DrawTop", "DrawWidth", "DrawHeight", Name, DataFilePath, DrawLeft, DrawTop, DrawWidth, DrawHeight)
    
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load the map's parameters from a string as stored in the GDP file."
    
    GeneralDeserialize Data, "Name", "DataFilePath", "DrawLeft", "DrawTop", "DrawWidth", "DrawHeight", Name, DataFilePath, DrawLeft, DrawTop, DrawWidth, DrawHeight
    
End Sub

Public Function Save(ByVal FileName As String)
Attribute Save.VB_Description = "Write a map file containing all the binary map-related data."
    Dim I As Integer
    Dim J As Integer
    Dim SLen As Long
    Dim SPos As Long
    Dim FF As Integer
    Dim S As String
    Dim sngVal As Single
    Dim iVal As Integer
    Dim lVal As Long
    Dim BA() As Byte
    
    On Error GoTo SaveErr
    FF = FreeFile
    Open FileName For Output As #FF
    Close #FF
    Open FileName For Binary As #FF
    S = "BMGC"
    Put #FF, , S
    S = "0002" ' File Version
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4 ' Reserve for length
    
    Put #FF, , PixWidth
    Put #FF, , PixHeight
    SaveString FF, BackgroundMusic
    If IsEmpty(Layers) Then
        I = 0
    Else
        I = UBound(Layers) - LBound(Layers) + 1
    End If
    Put #FF, , I ' Write layer count
    
    ' Now write the section data length
    SLen = Seek(FF) - SPos - 4 ' Exclude length of number
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    If Not IsEmpty(Layers) Then
        S = "LYRS" ' Layers section
        Put #FF, , S
        SPos = Seek(FF)
        Seek #FF, SPos + 4

        For I = LBound(Layers) To UBound(Layers)
            S = Layers(I).Name
            SLen = Len(S)
            Put #FF, , SLen
            Put #FF, , S
            
            S = Layers(I).TSDef.Name
            SLen = Len(S)
            Put #FF, , SLen
            Put #FF, , S
            
            sngVal = Layers(I).XScrollRate
            Put #FF, , sngVal
            sngVal = Layers(I).YScrollRate
            Put #FF, , sngVal
            S = IIf(Layers(I).Transparent, Chr$(1), Chr$(0))
            Put #FF, , S
            
            iVal = Layers(I).Columns
            Put #FF, , iVal
            iVal = Layers(I).Rows
            Put #FF, , iVal
            BA = Layers(I).Data.MapData
            Put #FF, , BA
        Next I
        
        ' Now write the section data length
        SLen = Seek(FF) - SPos - 4 ' Exclude length of number
        Seek #FF, SPos
        Put #FF, , SLen
        Seek #FF, SPos + SLen + 4
    End If
    
    S = "PATH" ' Paths section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    I = PathCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(PathColl) To UBound(PathColl)
            S = PathColl(I).Name
            SLen = Len(S)
            Put #FF, , SLen
            Put #FF, , S
            
            S = PathColl(I).LayerName
            SLen = Len(S)
            Put #FF, , SLen
            Put #FF, , S
            
            iVal = PathColl(I).PointCount
            Put #FF, , iVal
            For J = 0 To PathColl(I).PointCount - 1
                lVal = PathColl(I).PointX(J)
                Put #FF, , lVal
                lVal = PathColl(I).PointY(J)
                Put #FF, , lVal
            Next J
        Next
    End If
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    S = "SPTP" ' Sprite Templates section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    I = SpriteTemplateCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(SprTplColl) To UBound(SprTplColl)
            SprTplColl(I).Save FF
        Next I
    End If
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    S = "SPRD" ' Sprite Definitions section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    I = SpriteDefCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(SpriteDefColl) To UBound(SpriteDefColl)
            SpriteDefColl(I).Save FF
        Next I
    End If
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    S = "COLD" ' Collision Definitions Section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    I = CollDefCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(CollDefColl) To UBound(CollDefColl)
            CollDefColl(I).Save FF
        Next I
    End If
    For I = LBound(CollClassNames) To UBound(CollClassNames)
        SaveString FF, CollClassNames(I)
    Next
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    S = "SPEC" ' Special Functions Section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    I = SpecialCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(SpecialColl) To UBound(SpecialColl)
            SpecialColl(I).Save FF
        Next I
    End If
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4
    
    S = "IACT" ' Interactions section
    Put #FF, , S
    SPos = Seek(FF)
    Seek #FF, SPos + 4
    SaveString FF, PlayerSpriteName
    I = InteractCount
    Put #FF, , I
    If I > 0 Then
        For I = LBound(InteractColl) To UBound(InteractColl)
            InteractColl(I).Save FF
        Next I
    End If
    SLen = Seek(FF) - SPos - 4 ' Calculate section length
    Seek #FF, SPos
    Put #FF, , SLen
    Seek #FF, SPos + SLen + 4

    Close #FF
    DataFilePath = FileName
    IsDirty = False
    Exit Function

SaveErr:
    Close #FF
    If Not (Disp Is Nothing) Then Disp.Close
    MsgBox "Error saving map: " & Err.Description, vbCritical
End Function

Sub Load(ByVal FileName As String)
Attribute Load.VB_Description = "Load a complete binary map file into memory."
    Dim I As Integer
    Dim J As Integer
    Dim SLen As Long
    Dim FF As Integer
    Dim S As String
    Dim BA() As Byte
    Dim LName As String
    Dim LTDName As String
    Dim LXR As Single
    Dim LYR As Single
    Dim LTrans As Byte
    Dim LyrCount As Integer
    Dim iVal As Integer
    Dim ColCount As Integer
    Dim XVal As Long
    Dim YVal As Long

    On Error GoTo LoadErr
    FF = FreeFile
    Open FileName For Binary Access Read As #FF
    S = Space$(4)
    Get #FF, , S
    While S <> "BMGC"
        Get #FF, , SLen
        If (SLen + Seek(FF) > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            Err.Raise vbObjectError, , "File """ & FileName & """ is not a valid map"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Wend
    Get #FF, , S
    LoadVersion = CInt(S)
    If LoadVersion > 2 Then
        MsgBox "Map file """ & FileName & """ version too new", vbExclamation
        Close #FF
        Exit Sub
    End If
    Get #FF, , SLen
    Get #FF, , PixWidth
    Get #FF, , PixHeight
    BackgroundMusic = LoadString(FF)
    Get #FF, , LyrCount
    If LyrCount > 0 Then
        
        Get #FF, , S
        Do While S <> "LYRS"
            Get #FF, , SLen
            If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
                Close #FF
                Err.Raise vbObjectError, , "File """ & FileName & """ is not a valid map"
                Exit Sub
            End If
            Seek #FF, Seek(FF) + SLen
            Get #FF, , S
        Loop
        Get #FF, , SLen

        For I = 0 To LyrCount - 1
            Get #FF, , SLen
            LName = Space$(SLen)
            Get #FF, , LName
            
            Get #FF, , SLen
            LTDName = Space$(SLen)
            Get #FF, , LTDName
            
            Get #FF, , LXR
            Get #FF, , LYR
            Get #FF, , LTrans
                                    
            AddLayer LName, LTDName, LXR, LYR, LTrans > 0
            
            Get #FF, , iVal
            If Layers(I).Columns <> iVal Then Err.Raise vbObjectError, , "Layer data inconsistent with map"
            Get #FF, , iVal
            If Layers(I).Rows <> iVal Then Err.Raise vbObjectError, , "Layer data inconsistent with map"
            
            ReDim BA(0 To Layers(I).Columns * Layers(I).Rows - 1) As Byte
            Get #FF, , BA
            Layers(I).Data.MapData = BA
        Next I
    End If
    PathColl = Empty
    Get #FF, , S
    Do While S <> "PATH"
        Get #FF, , SLen
        If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain path information", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim PathColl(0 To ColCount - 1) As Path
        For I = 0 To ColCount - 1
            Set PathColl(I) = New Path
            
            Get #FF, , SLen
            LName = Space$(SLen)
            Get #FF, , LName
            PathColl(I).Name = LName
            
            Get #FF, , SLen
            LName = Space$(SLen)
            Get #FF, , LName
            PathColl(I).LayerName = LName
    
            Get #FF, , LyrCount
            For J = 0 To LyrCount - 1
                Get #FF, , XVal
                Get #FF, , YVal
                PathColl(I).AddPoint XVal, YVal
            Next J
        Next I
    Else
        PathColl = Empty
    End If
    
    SprTplColl = Empty
    Get #FF, , S
    Do While S <> "SPTP"
        Get #FF, , SLen
        If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain sprite templates", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim SprTplColl(0 To ColCount - 1) As SpriteTemplate
        For I = 0 To ColCount - 1
            Set SprTplColl(I) = New SpriteTemplate
            SprTplColl(I).Load FF
        Next
    End If

    SpriteDefColl = Empty
    Get #FF, , S
    Do While S <> "SPRD"
        Get #FF, , SLen
        If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain sprite definitions", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim SpriteDefColl(0 To ColCount - 1) As SpriteDef
        For I = 0 To ColCount - 1
            Set SpriteDefColl(I) = New SpriteDef
            SpriteDefColl(I).Load FF
        Next
    End If
    
    CollDefColl = Empty
    Get #FF, , S
    Do While S <> "COLD"
        Get #FF, , SLen
        If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain collision definitions", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim CollDefColl(0 To ColCount - 1) As CollisionDef
        For I = 0 To ColCount - 1
            Set CollDefColl(I) = New CollisionDef
            CollDefColl(I).Load FF
        Next
    End If
    For I = LBound(CollClassNames) To UBound(CollClassNames)
        CollClassNames(I) = LoadString(FF)
    Next

    SpecialColl = Empty
    Get #FF, , S
    Do While S <> "SPEC"
        Get #FF, , SLen
        If (Seek(FF) + SLen > LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain Special Functions section", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim SpecialColl(0 To ColCount - 1) As SpecialFunction
        For I = 0 To ColCount - 1
            Set SpecialColl(I) = New SpecialFunction
            SpecialColl(I).Load FF
        Next
    End If

    InteractColl = Empty
    Get #FF, , S
    Do While S <> "IACT"
        Get #FF, , SLen
        If (Seek(FF) + SLen < LOF(FF)) Or (SLen < 0) Then
            Close #FF
            MsgBox "File """ & FileName & """ does not contain interaction definitions", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        Seek #FF, Seek(FF) + SLen
        Get #FF, , S
    Loop
    Get #FF, , SLen
    PlayerSpriteName = LoadString(FF)
    Get #FF, , ColCount
    If ColCount > 0 Then
        ReDim InteractColl(0 To ColCount - 1) As Interaction
        For I = 0 To ColCount - 1
            Set InteractColl(I) = New Interaction
            InteractColl(I).Load FF
        Next
    End If

    Close #FF
    DataFilePath = FileName
    IsDirty = False
    Exit Sub

LoadErr:
    Close #FF
    If Not (Disp Is Nothing) Then Disp.Close
    MsgBox "Error loading map: " & Err.Description, vbCritical
End Sub

Sub Draw(ByVal XOffset As Long, ByVal YOffset As Long, Optional ByVal bAdvanceTileAnims As Boolean = True)
Attribute Draw.VB_Description = "Draw all layer tiles and sprites according to the scroll offset and map view position."
    Dim I As Integer

    If IsEmpty(Layers) Then Exit Sub
    
    bActive = True
    For I = LBound(Layers) To UBound(Layers)
        With Layers(I)
            If Not .IsLoaded Then
                On Error Resume Next
                Disp.Close
                On Error GoTo 0
                MsgBox "Layer " & CStr(I) & " is not loaded; Can't draw"
                Exit Sub
            End If
            If bAdvanceTileAnims Then .AdvanceTileAnims
            Disp.DrawMap .Tileset, .Data, CLng(XOffset * .XScrollRate), CLng(YOffset * .YScrollRate), DrawLeft, DrawTop, DrawWidth, DrawHeight, .Transparent
            .DrawLayerSprites XOffset, YOffset
        End With
    Next
End Sub

Property Get ViewWidth() As Integer
Attribute ViewWidth.VB_Description = "Get/Set the width of the map window."
    ViewWidth = DrawWidth
End Property

Property Get ViewHeight() As Integer
Attribute ViewHeight.VB_Description = "Get/Set the height of the map window."
    ViewHeight = DrawHeight
End Property

Property Get ViewLeft() As Integer
Attribute ViewLeft.VB_Description = "Get/Set the position of the map window."
    ViewLeft = DrawLeft
End Property

Property Get ViewTop() As Integer
Attribute ViewTop.VB_Description = "Get/Set the position of the map window."
    ViewTop = DrawTop
End Property

Property Let ViewWidth(ByVal NewVal As Integer)
    DrawWidth = NewVal
    IsDirty = True
End Property

Property Let ViewHeight(ByVal NewVal As Integer)
    DrawHeight = NewVal
    IsDirty = True
End Property

Property Let ViewLeft(ByVal NewVal As Integer)
    DrawLeft = NewVal
    IsDirty = True
End Property

Property Let ViewTop(ByVal NewVal As Integer)
    DrawTop = NewVal
    IsDirty = True
End Property

Property Get MapWidth() As Long
Attribute MapWidth.VB_Description = "Get or set the width of the total map area in pixels (doesn't update layers)."
Attribute MapWidth.VB_ProcData.VB_Invoke_Property = ";Data"
    MapWidth = PixWidth
End Property

Property Let MapWidth(ByVal NewVal As Long)
    PixWidth = NewVal
    IsDirty = True
End Property

Property Get MapHeight() As Long
Attribute MapHeight.VB_Description = "Get or set the height of the total map area in pixels (doesn't update layers)."
Attribute MapHeight.VB_ProcData.VB_Invoke_Property = ";Data"
    MapHeight = PixHeight
End Property

Property Let MapHeight(ByVal NewVal As Long)
    PixHeight = NewVal
    IsDirty = True
End Property

Property Get MapLayer(ByVal Index As Variant) As Layer
Attribute MapLayer.VB_Description = "Return the layer by it's ordinal index in the map."
Attribute MapLayer.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute MapLayer.VB_UserMemId = 0
    Dim I As Integer
    
    If VarType(Index) = vbString Then
        For I = LBound(Layers) To UBound(Layers)
            If UCase(Layers(I).Name) = UCase(Index) Then Exit For
        Next
    Else
        I = CInt(Index)
    End If

    If Not IsEmpty(Layers) Then
        Set MapLayer = Layers(I)
    Else
        Set MapLayer = Nothing
    End If
    
End Property

Public Property Get LayerExists(ByVal Key As Variant) As Boolean
Attribute LayerExists.VB_Description = "Return True if the specified layer name or index exists in the map."
    Dim I As Integer
    
    If IsEmpty(Layers) Then
        LayerExists = False
        Exit Property
    End If
    
    If VarType(Key) = vbString Then
        For I = LBound(Layers) To UBound(Layers)
            If UCase(Layers(I).Name) = UCase(Key) Then Exit For
        Next
    Else
        I = CInt(Key)
    End If

    If I < 0 Or I > UBound(Layers) Then
        LayerExists = False
    Else
        LayerExists = True
    End If

End Property

Public Sub ChangeScreenTile(ByVal XOff As Long, ByVal YOff As Long, ByVal X As Integer, ByVal Y As Integer, ByVal LayerIndex As Integer, ByVal TileVal As Byte)
Attribute ChangeScreenTile.VB_Description = "Change the tile on a particular layer beneath a particular coordinate on the screen given the map's scroll position."
    Dim Row As Integer, Col As Integer
    
    If IsEmpty(Layers) Then Exit Sub
    If LayerIndex < LBound(Layers) Or LayerIndex > UBound(Layers) Then
        On Error Resume Next
        Disp.Close
        MsgBox "Layer index out of bounds", vbCritical
        Exit Sub
    End If
    
    If ScreenToTile(XOff, YOff, X, Y, LayerIndex, Col, Row) Then
        Layers(LayerIndex).Data.TileValue(Col, Row) = TileVal
    End If
    IsDirty = True
    
End Sub

Public Function ScreenToTile(ByVal XOff As Long, ByVal YOff As Long, ByVal X As Integer, ByVal Y As Integer, ByVal LayerIndex As Integer, ByRef Col As Integer, ByRef Row As Integer) As Boolean
Attribute ScreenToTile.VB_Description = "Calculate a tile's row and column in the layer from a pixel coordinate on the screen and the scroll position of the map."
    
    If IsEmpty(Layers) Then Exit Function
    If LayerIndex < LBound(Layers) Or LayerIndex > UBound(Layers) Then
        On Error Resume Next
        Disp.Close
        MsgBox "Layer index out of bounds", vbCritical
        Exit Function
    End If

    With Layers(LayerIndex)
        Col = (XOff * .XScrollRate + X - ViewLeft) \ .TSDef.TileWidth
        Row = (YOff * .YScrollRate + Y - ViewTop) \ .TSDef.TileHeight
    End With
    
    If X < ViewLeft Or X >= ViewLeft + ViewWidth Or Y < ViewTop Or Y > ViewTop + ViewHeight Then
        ScreenToTile = False
    Else
        ScreenToTile = True
    End If
    
End Function

Property Get LayerCount() As Integer
Attribute LayerCount.VB_Description = "Return the number of layers in the map."
    If IsEmpty(Layers) Then
        LayerCount = 0
    Else
        LayerCount = UBound(Layers) - LBound(Layers) + 1
    End If
End Property

Public Property Get Path() As String
Attribute Path.VB_Description = "The path of the binary map file for this map."
    Path = DataFilePath
End Property

Public Property Let Path(ByVal NewVal As String)
    DataFilePath = NewVal
    IsDirty = True
End Property

Public Sub ShiftLayer(ByVal Index As Integer, ByVal Dir As Integer)
Attribute ShiftLayer.VB_Description = "Move a layer in front of or behind other layers. (Positive Dir = foreground.)"
    Dim Lr As Layer
    Dim Idx As Integer
    
    Set Lr = Layers(Index)
    Set Layers(Index) = Layers(Index + Dir)
    Set Layers(Index + Dir) = Lr
    
    For Idx = 0 To Lr.pMap.SpecialCount - 1
        With Lr.pMap
            If .Specials(Idx).LayerIndex = Index Then
                .Specials(Idx).LayerIndex = Index + Dir
            ElseIf .Specials(Idx).LayerIndex = Index + Dir Then
                .Specials(Idx).LayerIndex = Index
            End If
        End With
    Next
    
    Set Lr = Nothing
    
    IsDirty = True
End Sub

Public Sub LoadTileAnims()
Attribute LoadTileAnims.VB_Description = "Initialize tile animation (AnimDef) on all layers of the map."
    Dim I As Integer
    
    If IsEmpty(Layers) Then Exit Sub
    For I = LBound(Layers) To UBound(Layers)
        Layers(I).LoadTileAnims
    Next
    
End Sub

Public Sub AddSpriteDef(SD As SpriteDef)
Attribute AddSpriteDef.VB_Description = "Add a new SpriteDef to the map."
    If IsEmpty(SpriteDefColl) Then
        ReDim SpriteDefColl(0) As SpriteDef
    Else
        ReDim Preserve SpriteDefColl(LBound(SpriteDefColl) To UBound(SpriteDefColl) + 1)
    End If
    Set SpriteDefColl(UBound(SpriteDefColl)) = SD
    IsDirty = True
End Sub

Public Sub RemoveSpriteDef(ByVal Key As Variant)
Attribute RemoveSpriteDef.VB_Description = "Remove a SpriteDef object from the map."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(SpriteDefColl) To UBound(SpriteDefColl)
            If SpriteDefColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(SpriteDefColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveSpriteDef"
    End If
    
    SpriteDefColl(Index).ReleaseRefs
    
    Do While Index < UBound(SpriteDefColl)
        Set SpriteDefColl(Index) = SpriteDefColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set SpriteDefColl(UBound(SpriteDefColl)) = Nothing
    
    If Index > LBound(SpriteDefColl) Then
        ReDim Preserve SpriteDefColl(LBound(SpriteDefColl) To UBound(SpriteDefColl) - 1)
    Else
        SpriteDefColl = Empty
    End If
    IsDirty = True
End Sub

Property Set SpriteDefs(ByVal Key As Variant, NewDef As SpriteDef)
Attribute SpriteDefs.VB_Description = "Return a SpriteDef object given its name or ordinal Index in the map."
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(SpriteDefColl) To UBound(SpriteDefColl)
            If UCase$(SpriteDefColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(SpriteDefColl) Then
        Err.Raise vbObjectError, , "SpriteDef " & CStr(Key) & " not found"
    End If
    
    Set SpriteDefColl(I) = NewDef
End Property

Property Get SpriteDefs(ByVal Key As Variant) As SpriteDef
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(SpriteDefColl) To UBound(SpriteDefColl)
            If UCase$(SpriteDefColl(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(SpriteDefColl) Then
        Err.Raise vbObjectError, , "SpriteDef " & CStr(Key) & " not found"
    End If
    
    Set SpriteDefs = SpriteDefColl(I)
End Property

Public Property Get SpriteDefCount() As Integer
Attribute SpriteDefCount.VB_Description = "Return the number of SpriteDef objects in the map."
    If IsEmpty(SpriteDefColl) Then
        SpriteDefCount = 0
    Else
        SpriteDefCount = UBound(SpriteDefColl) - LBound(SpriteDefColl) + 1
    End If
End Property

Public Sub AddSpriteTemplate(Template As SpriteTemplate)
Attribute AddSpriteTemplate.VB_Description = "Add a new SpriteTemplate to the map."
    If IsEmpty(SprTplColl) Then
        ReDim SprTplColl(0) As SpriteTemplate
    Else
        ReDim Preserve SprTplColl(LBound(SprTplColl) To UBound(SprTplColl) + 1)
    End If
    Set SprTplColl(UBound(SprTplColl)) = Template
    IsDirty = True
End Sub

Public Sub RemoveSpriteTemplate(ByVal Key As Variant)
Attribute RemoveSpriteTemplate.VB_Description = "Remove a SpriteTemplate object from the map."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(SprTplColl) To UBound(SprTplColl)
            If SprTplColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(SprTplColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveSpriteTemplate"
    End If
    
    Do While Index < UBound(SprTplColl)
        Set SprTplColl(Index) = SprTplColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set SprTplColl(UBound(SprTplColl)) = Nothing
    
    If Index > LBound(SprTplColl) Then
        ReDim Preserve SprTplColl(LBound(SprTplColl) To UBound(SprTplColl) - 1)
    Else
        SprTplColl = Empty
    End If
    IsDirty = True
End Sub

Public Property Get SpriteTemplateExists(ByVal Key As Variant) As Boolean
Attribute SpriteTemplateExists.VB_Description = "Return True if the SpriteTemplate name or index exists in the Map."
    Dim Index As Integer
    
    If IsEmpty(SprTplColl) Then
        SpriteTemplateExists = False
        Exit Property
    End If
    
    If VarType(Key) = vbString Then
        For Index = LBound(SprTplColl) To UBound(SprTplColl)
            If SprTplColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If

    If Index < 0 Or Index > UBound(SprTplColl) Then
        SpriteTemplateExists = False
    Else
        SpriteTemplateExists = True
    End If
End Property

Public Property Get SpriteTemplates(ByVal Key As Variant) As SpriteTemplate
Attribute SpriteTemplates.VB_Description = "Return a SpriteTemplate given its name or ordinal index in the Map."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(SprTplColl) To UBound(SprTplColl)
            If SprTplColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(SprTplColl) Then
        Err.Raise vbObjectError, , "Invalid key for SpriteTemplates"
    End If
    
    Set SpriteTemplates = SprTplColl(Index)
End Property

Public Property Set SpriteTemplates(ByVal Key As Variant, NewVal As SpriteTemplate)
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(SprTplColl) To UBound(SprTplColl)
            If SprTplColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(SprTplColl) Then
        Err.Raise vbObjectError, , "Invalid key for SpriteTemplates"
    End If
    
    Set SprTplColl(Index) = NewVal
End Property

Public Property Get SpriteTemplateCount() As Integer
Attribute SpriteTemplateCount.VB_Description = "Return the number of SpriteTemplate objects in the map."
    If IsEmpty(SprTplColl) Then
        SpriteTemplateCount = 0
    Else
        SpriteTemplateCount = UBound(SprTplColl) - LBound(SprTplColl) + 1
    End If
End Property

Public Sub InitSprites()
Attribute InitSprites.VB_Description = "Initialize all this map's sprites, creating Sprites from SpriteDefs."
    Dim I As Integer
   
    For I = 0 To LayerCount - 1
        Layers(I).ClearSprites
    Next
    For I = 0 To SpriteDefCount() - 1
        If SpriteDefColl(I).Flags And eDefFlagBits.FLAG_INSTANCE Then SpriteDefColl(I).rLayer.AddSprite SpriteDefColl(I).MakeInstance
    Next
   
End Sub

Public Sub AdvanceSprites()
Attribute AdvanceSprites.VB_Description = "Move all sprites on all layers of this map according to their movement definitions."
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim ClsMsk As Integer
    Static Msk(0 To 15) As Integer
    Dim LyrWid As Long, LyrHgt As Long
    Dim LYR As Layer
    Dim SpLast As Integer
    
    If Msk(0) = 0 Then
        For I = 0 To 14
            Msk(I) = 2 ^ I
        Next
        Msk(15) = -32768
    End If
    
    For I = 0 To LayerCount - 1
        Set LYR = Layers(I)
        With LYR
            LyrWid = .Columns * .TSDef.TileWidth
            LyrHgt = .Rows * .TSDef.TileHeight
            For J = 0 To 15
                .ClsMbrCnt(J) = 0
            Next
            SpLast = .SpriteCount - 1
            ' First advance all potential platforms, then advance riders
            For J = 0 To SpLast
                With .Sprite(J)
                    If .rDef.Template.ControlType = CONTROL_SIMPLE Then Exit For
                    If .pRideOnRef Is Nothing Then
                        .Advance
                    End If
                    ClsMsk = .rDef.Template.CollClass
                End With
                For K = 0 To 15
                    If Msk(K) And ClsMsk Then .ClsMbrCnt(K) = .ClsMbrCnt(K) + 1
                Next
            Next
            Do While J < SpLast
                With .Sprite(J)
                    .X = .X + .DX: .Y = .Y + .DY
                    .AdvanceFrame 1
                    If .X > LyrWid Or .Y > LyrHgt Or .X < 0 Or .Y < 0 Then
                        LYR.RemoveSprite J
                        SpLast = SpLast - 1
                    End If
                End With
                J = J + 1
            Loop
            For J = 0 To SpLast
                With .Sprite(J)
                    If .rDef.Template.ControlType = CONTROL_SIMPLE Then Exit For
                    If Not (.pRideOnRef Is Nothing) Then
                        .Advance
                    End If
                End With
            Next
        End With
    Next
    
End Sub

Public Sub CheckSpriteCollisions()
Attribute CheckSpriteCollisions.VB_Description = "Check for and process all sprite collisions on all layers of the map."
    Dim I As Integer
    
    If IsEmpty(CollDefColl) Then Exit Sub
    For I = 0 To LayerCount - 1
        Layers(I).CheckSpriteCollisions
    Next
End Sub

Public Property Get Interactions(ByVal Index As Integer) As Interaction
Attribute Interactions.VB_Description = "Return or set an interaction object in the map according to it's ordinal index."
    Set Interactions = InteractColl(Index)
End Property

Public Property Set Interactions(ByVal Index As Integer, NewVal As Interaction)
    Set InteractColl(Index) = NewVal
End Property

Public Property Get InteractCount() As Integer
Attribute InteractCount.VB_Description = "Return the number of Interaction definitions for this map."
    If IsEmpty(InteractColl) Then
        InteractCount = 0
    Else
        InteractCount = UBound(InteractColl) - LBound(InteractColl) + 1
    End If
End Property

Public Sub AddInteraction(InteractionObj As Interaction)
Attribute AddInteraction.VB_Description = "Add a new Interaction to the map."
    If IsEmpty(InteractColl) Then
        ReDim InteractColl(0) As Interaction
    Else
        ReDim Preserve InteractColl(LBound(InteractColl) To UBound(InteractColl) + 1) As Interaction
    End If
    Set InteractColl(UBound(InteractColl)) = InteractionObj
End Sub

Public Sub RemoveInteraction(ByVal Index As Integer)
Attribute RemoveInteraction.VB_Description = "Remove an Interaction object from the map"
    Dim Idx As Integer
    
    For Idx = Index To UBound(InteractColl) - 1
        Set InteractColl(Idx) = InteractColl(Idx + 1)
    Next
    Set InteractColl(UBound(InteractColl)) = Nothing
    If UBound(InteractColl) >= 1 Then
        ReDim Preserve InteractColl(LBound(InteractColl) To UBound(InteractColl) - 1)
    Else
        InteractColl = Empty
    End If
End Sub

