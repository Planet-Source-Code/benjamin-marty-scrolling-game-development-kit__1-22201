VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Through this object, all aspects of a project (all maps & graphics etc) are accessible.  Get instance with Engine.ProjectObj."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Project.cls - Game Project Class
'
'======================================================================

Option Explicit

Dim TileSetDefs As Variant
Dim MapColl As Variant
Dim MatchColl As Variant
Dim AnimColl As Variant
Dim GroupColl As Variant
Dim SolidColl As Variant
Public ProjectPath As String
Attribute ProjectPath.VB_VarUserMemId = -518
Attribute ProjectPath.VB_VarProcData = ";Text"
Attribute ProjectPath.VB_VarDescription = "Return or set the path to the GDP plain text project file."
Public IsDirty As Boolean
Attribute IsDirty.VB_VarDescription = "Returns True if the project data (strictly that stored in the GDP) has changed."
Public GamePlayer As Player
Attribute GamePlayer.VB_VarDescription = "Returns the single Player object for the project."
Public MediaMgr As MediaManager
Attribute MediaMgr.VB_VarDescription = "Return the single MediaManager object for the current project."
Public bSplashShowing As Boolean
Attribute bSplashShowing.VB_VarDescription = "Returns True if the spash screen is still showing from initial startup."

Event OnEditMap(Mp As MapEdit)
Event OnEditTileset(TE As TileEdit)
Event GenericEvent(Data As Variant)

Public Sub TriggerEditMap(Mp As MapEdit)
    RaiseEvent OnEditMap(Mp)
End Sub

Public Sub TriggerTileEdit(TE As TileEdit)
    RaiseEvent OnEditTileset(TE)
End Sub

Public Sub TriggerGeneric(Data As Variant)
    RaiseEvent GenericEvent(Data)
End Sub

Public Sub WaitForSplash()
Attribute WaitForSplash.VB_Description = "Wait for the splash screen to disappear, then return."
    Do While bSplashShowing
        DoEvents
    Loop
End Sub

Public Function AddTileSet(ByVal ImagePath As String, ByVal TileWidth As Integer, ByVal TileHeight As Integer, ByVal Name As String) As TileSetDef
Attribute AddTileSet.VB_Description = "Add a new TileSetDef to the project."
    If IsEmpty(TileSetDefs) Then
        ReDim TileSetDefs(0) As TileSetDef
    Else
        ReDim Preserve TileSetDefs(LBound(TileSetDefs) To UBound(TileSetDefs) + 1)
    End If
    Set TileSetDefs(UBound(TileSetDefs)) = New TileSetDef
    With TileSetDefs(UBound(TileSetDefs))
        If ImagePath <> "" Then
            .ImagePath = GetRelativePath(ProjectPath, ImagePath)
        End If
        .TileWidth = TileWidth
        .TileHeight = TileHeight
        .Name = Name
    End With
    Set AddTileSet = TileSetDefs(UBound(TileSetDefs))
    IsDirty = True
End Function

Public Sub AddMap(M As Map)
Attribute AddMap.VB_Description = "Add a new Map to the project"
    If IsEmpty(MapColl) Then
        ReDim MapColl(0) As Map
    Else
        ReDim Preserve MapColl(LBound(MapColl) To UBound(MapColl) + 1)
    End If
    Set MapColl(UBound(MapColl)) = M
    IsDirty = True
End Sub

Public Sub AddMatch(M As MatchDef)
Attribute AddMatch.VB_Description = "Add a new MatchDef to the project"
    If IsEmpty(MatchColl) Then
        ReDim MatchColl(0) As MatchDef
    Else
        ReDim Preserve MatchColl(LBound(MatchColl) To UBound(MatchColl) + 1)
    End If
    Set MatchColl(UBound(MatchColl)) = M
    IsDirty = True
End Sub

Public Sub AddAnim(A As AnimDef)
Attribute AddAnim.VB_Description = "Add a new AnimDef to the project."
    If IsEmpty(AnimColl) Then
        ReDim AnimColl(0) As AnimDef
    Else
        ReDim Preserve AnimColl(LBound(AnimColl) To UBound(AnimColl) + 1)
    End If
    Set AnimColl(UBound(AnimColl)) = A
    IsDirty = True
End Sub

Public Function AddGroup(ByVal GroupName As String, ByVal TilesetName As String) As Category
Attribute AddGroup.VB_Description = "Add a new Category to the project."
    If IsEmpty(GroupColl) Then
        ReDim GroupColl(0) As Category
    Else
        ReDim Preserve GroupColl(LBound(GroupColl) To UBound(GroupColl) + 1)
    End If
    Set GroupColl(UBound(GroupColl)) = New Category
    With GroupColl(UBound(GroupColl))
        .Name = GroupName
        .TSName = TilesetName
    End With
    Set AddGroup = GroupColl(UBound(GroupColl))
    IsDirty = True
End Function

Public Function AddSolidDef(ByVal SolidDefName As String, ByVal TilesetName As String) As SolidDef
Attribute AddSolidDef.VB_Description = "Add a new SolidDef to the project."
    If IsEmpty(SolidColl) Then
        ReDim SolidColl(0) As SolidDef
    Else
        ReDim Preserve SolidColl(LBound(SolidColl) To UBound(SolidColl) + 1)
    End If
    Set SolidColl(UBound(SolidColl)) = New SolidDef
    With SolidColl(UBound(SolidColl))
        .Name = SolidDefName
        .TSName = TilesetName
    End With
    Set AddSolidDef = SolidColl(UBound(SolidColl))
    IsDirty = True
End Function

Public Sub RemoveMap(ByVal Key As Variant)
Attribute RemoveMap.VB_Description = "Remove a Map given its name or ordinal position in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(MapColl) To UBound(MapColl)
            If MapColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(MapColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveMap"
    End If
    
    Do While Index < UBound(MapColl)
        Set MapColl(Index) = MapColl(Index + 1)
        Index = Index + 1
    Loop
    Set MapColl(UBound(MapColl)) = Nothing
    
    If Index > LBound(MapColl) Then
        ReDim Preserve MapColl(LBound(MapColl) To UBound(MapColl) - 1)
    Else
        MapColl = Empty
    End If
    IsDirty = True
    
End Sub

Public Sub RemoveMatch(ByVal Key As Variant)
Attribute RemoveMatch.VB_Description = "Remove a MatchDef given its name or ordinal position in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(MatchColl) To UBound(MatchColl)
            If MatchColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(MatchColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveMatch"
    End If
    
    Do While Index < UBound(MatchColl)
        Set MatchColl(Index) = MatchColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set MatchColl(UBound(MatchColl)) = Nothing
    
    If Index > LBound(MatchColl) Then
        ReDim Preserve MatchColl(LBound(MatchColl) To UBound(MatchColl) - 1)
    Else
        MatchColl = Empty
    End If
    IsDirty = True
    
End Sub

Public Sub RemoveAnim(ByVal Key As Variant)
Attribute RemoveAnim.VB_Description = "Remove an AnimDef given its name or ordinal position in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(AnimColl) To UBound(AnimColl)
            If AnimColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(AnimColl) Then
        Err.Raise vbObjectError, , "Invalid key for RemoveAnim"
    End If
    
    Do While Index < UBound(AnimColl)
        Set AnimColl(Index) = AnimColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set AnimColl(UBound(AnimColl)) = Nothing
    
    If Index > LBound(AnimColl) Then
        ReDim Preserve AnimColl(LBound(AnimColl) To UBound(AnimColl) - 1)
    Else
        AnimColl = Empty
    End If
    IsDirty = True
    
End Sub

Public Sub RemoveGroup(ByVal GroupName As String, ByVal TilesetName As String)
Attribute RemoveGroup.VB_Description = "Remove a Category from the project given its name and Tileset name."
    Dim Index As Integer
    
    For Index = LBound(GroupColl) To UBound(GroupColl)
        If GroupColl(Index).Name = GroupName And GroupColl(Index).TSName = TilesetName Then
            Exit For
        End If
    Next
    
    If Index > UBound(GroupColl) Then
        Err.Raise vbObjectError, , "Group " & GroupName & " for tileset " & TilesetName & " not found."
    End If
    
    Set GroupColl(Index) = GroupColl(UBound(GroupColl))
    Set GroupColl(UBound(GroupColl)) = Nothing
    If UBound(GroupColl) - LBound(GroupColl) <= 0 Then
        GroupColl = Empty
    Else
        ReDim Preserve GroupColl(LBound(GroupColl) To UBound(GroupColl) - 1)
    End If
    IsDirty = True
End Sub

Public Sub RemoveSolidDef(Item As SolidDef)
Attribute RemoveSolidDef.VB_Description = "Remove a SolidDef from the project given a reference to it."
    Dim Index As Integer
    
    For Index = LBound(SolidColl) To UBound(SolidColl)
        If SolidColl(Index) Is Item Then Exit For
    Next
    
    If Index > UBound(SolidColl) Then
        Err.Raise vbObjectError, , "RemoveSolidDef cannot remove given item"
    End If
    
    Do While Index < UBound(SolidColl)
        Set SolidColl(Index) = SolidColl(Index + 1)
        Index = Index + 1
    Loop
    
    Set SolidColl(Index) = Nothing
    
    If Index > LBound(SolidColl) Then
        ReDim Preserve SolidColl(LBound(SolidColl) To UBound(SolidColl) - 1)
    Else
        SolidColl = Empty
    End If
    IsDirty = True
    
End Sub

Public Property Get Maps(ByVal Key As Variant) As Map
Attribute Maps.VB_Description = "Return a map given its name or ordinal index in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(MapColl) To UBound(MapColl)
            If MapColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = Key
    End If
    
    If Index < 0 Or Index > UBound(MapColl) Then
        Err.Raise vbObjectError, , "Invalid key for Map"
    End If
    
    Set Maps = MapColl(Index)
End Property

Public Property Get MapExists(ByVal Key As Variant) As Boolean
Attribute MapExists.VB_Description = "Return True if a map by the specified name or ordinal index exists."
    Dim Index As Integer
    
    If IsEmpty(MapColl) Then
        MapExists = False
        Exit Property
    End If
    
    If VarType(Key) = vbString Then
        For Index = LBound(MapColl) To UBound(MapColl)
            If MapColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = Key
    End If
    
    If Index < 0 Or Index > UBound(MapColl) Then
        MapExists = False
    Else
        MapExists = True
    End If
    
End Property

Public Property Get MatchDefs(ByVal Key As Variant) As MatchDef
Attribute MatchDefs.VB_Description = "Return a MatchDef object given its name or ordinal index in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(MatchColl) To UBound(MatchColl)
            If MatchColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(MatchColl) Then
        Err.Raise vbObjectError, , "Invalid key for MatchDef"
    End If

    Set MatchDefs = MatchColl(Index)
    
End Property

Public Property Get AnimDefs(ByVal Key As Variant) As AnimDef
Attribute AnimDefs.VB_Description = "Return an AnimDef based on its name or 0-based index in the project."
    Dim Index As Integer
    
    If VarType(Key) = vbString Then
        For Index = LBound(AnimColl) To UBound(AnimColl)
            If AnimColl(Index).Name = Key Then Exit For
        Next
    Else
        Index = CInt(Key)
    End If
    
    If Index < 0 Or Index > UBound(AnimColl) Then
        Err.Raise vbObjectError, , "Invalid key for AnimDef"
    End If
    
    Set AnimDefs = AnimColl(Index)
End Property

Public Function Groups(ByVal GroupName As String, ByVal TilesetName As String) As Category
Attribute Groups.VB_Description = "Returns the Category object by the specified name for the specified tileset."
    Dim Index As Integer
    
    For Index = LBound(GroupColl) To UBound(GroupColl)
        If GroupColl(Index).Name = GroupName And GroupColl(Index).TSName = TilesetName Then
            Exit For
        End If
    Next
    
    If Index > UBound(GroupColl) Then
        Err.Raise vbObjectError, , "Group " & GroupName & " for tileset " & TilesetName & " not found."
    End If
    
    Set Groups = GroupColl(Index)
    
End Function

Public Function GroupByIndex(ByVal Index As Integer) As Category
Attribute GroupByIndex.VB_Description = "Return a Category object by ordinal position in the list of all Category objects in the project."
    Set GroupByIndex = GroupColl(Index)
End Function

Public Function GroupCount() As Integer
Attribute GroupCount.VB_Description = "Return the total number of Category objects in the project."
    If IsEmpty(GroupColl) Then
        GroupCount = 0
        Exit Function
    End If
    GroupCount = UBound(GroupColl) - LBound(GroupColl) + 1
End Function

Public Function TilesetGroupByIndex(ByVal TileSetKey As String, ByVal Index As Integer) As Category
Attribute TilesetGroupByIndex.VB_Description = "Find and return the Nth Category for the specified TileSetDef where N=Index."
    Dim I As Integer
    Dim Cnt As Integer
    
    Cnt = 0
    For I = LBound(GroupColl) To UBound(GroupColl)
        If GroupColl(I).TSName = TileSetKey Then
            If Cnt = Index Then
                Set TilesetGroupByIndex = GroupColl(I)
                Exit Function
            End If
            Cnt = Cnt + 1
        End If
    Next
    Err.Raise vbObjectError, , "Could not find group #" & Index & " for Tileset " & TileSetKey
End Function

Public Function GroupByTilesetCount(ByVal TileSetKey As String) As Integer
Attribute GroupByTilesetCount.VB_Description = "Return the number of Category objects in the project for the specified tileset name."
    Dim Index As Integer
    Dim Cnt As Integer
    
    If IsEmpty(GroupColl) Then
        GroupByTilesetCount = 0
        Exit Function
    End If
    
    Cnt = 0
    For Index = LBound(GroupColl) To UBound(GroupColl)
        If GroupColl(Index).TSName = TileSetKey Then
            Cnt = Cnt + 1
        End If
    Next
    GroupByTilesetCount = Cnt
End Function

Public Function GroupExists(ByVal GroupName As String, ByVal TilesetName As String) As Boolean
Attribute GroupExists.VB_Description = "Return True if a Category by the specified name exists for the specified tileset."
    Dim Index As Integer
    
    If IsEmpty(GroupColl) Then
        GroupExists = False
        Exit Function
    End If
    
    For Index = LBound(GroupColl) To UBound(GroupColl)
        If GroupColl(Index).Name = GroupName And GroupColl(Index).TSName = TilesetName Then
            Exit For
        End If
    Next
    
    If Index > UBound(GroupColl) Then
        GroupExists = False
    Else
        GroupExists = True
    End If
    
End Function

Public Function SolidDefs(ByVal SolidDefName As String, ByVal TilesetName As String) As SolidDef
Attribute SolidDefs.VB_Description = "Return the named SolidDef object for the specified TileSetDef."
    Dim Index As Integer
    
    For Index = LBound(SolidColl) To UBound(SolidColl)
        If SolidColl(Index).Name = SolidDefName And SolidColl(Index).TSName = TilesetName Then Exit For
    Next
    
    If Index > UBound(SolidColl) Then
        Err.Raise vbObjectError, , "SolidDef(" & TilesetName & ", " & SolidDefName & ") not found in SolidDefs"
    End If
    
    Set SolidDefs = SolidColl(Index)
End Function

Public Function SolidDefsByIndex(ByVal Index As Integer) As SolidDef
Attribute SolidDefsByIndex.VB_Description = "Return a SolidDef object based on its ordinal index of all SolidDefs in the project."
    Set SolidDefsByIndex = SolidColl(Index)
End Function

Public Function SolidDefCount() As Integer
Attribute SolidDefCount.VB_Description = "Return the total number of SolidDef objects in the project."
    If IsEmpty(SolidColl) Then
        SolidDefCount = 0
    Else
        SolidDefCount = UBound(SolidColl) - LBound(SolidColl) + 1
    End If
End Function

Public Function SolidDefExists(ByVal SolidDefName As String, ByVal TilesetName As String) As Boolean
Attribute SolidDefExists.VB_Description = "Return true if the specified SolidDef exists for the named tileset."
    Dim Index As Integer
    
    If IsEmpty(SolidColl) Then
        SolidDefExists = False
        Exit Function
    End If
    
    For Index = LBound(SolidColl) To UBound(SolidColl)
        If SolidColl(Index).Name = SolidDefName And SolidColl(Index).TSName = TilesetName Then Exit For
    Next
        
    If Index > UBound(SolidColl) Then
        SolidDefExists = False
    Else
        SolidDefExists = True
    End If
End Function

Public Function SolidDefIndexByTileset(ByVal TilesetName As String, ByVal Index As Integer) As Integer
Attribute SolidDefIndexByTileset.VB_Description = "Return the Nth SolidDef object for the named tileset where N=Index."
    Dim I As Integer
    Dim Cnt As Integer
    
    For I = LBound(SolidColl) To UBound(SolidColl)
        If SolidColl(I).TSName = TilesetName Then
            If Cnt = Index Then
                SolidDefIndexByTileset = I
                Exit Function
            End If
            Cnt = Cnt + 1
        End If
    Next
    
    Err.Raise vbObjectError, , "SolidDefIndexByTileset index out of bounds"
    
End Function

Public Function SolidDefByTilesetCount(ByVal TilesetName As String) As Integer
Attribute SolidDefByTilesetCount.VB_Description = "Return the number of SolidDefs for the specified Tileset."
    Dim I As Integer
    Dim Cnt As Integer
    
    If IsEmpty(SolidColl) Then
        SolidDefByTilesetCount = 0
        Exit Function
    End If
    
    For I = LBound(SolidColl) To UBound(SolidColl)
        If SolidColl(I).TSName = TilesetName Then
            Cnt = Cnt + 1
        End If
    Next
    
    SolidDefByTilesetCount = Cnt
    
End Function

Public Property Get TileSetDef(ByVal Key As Variant) As TileSetDef
Attribute TileSetDef.VB_Description = "Return a TileSetDef given its name or ordinal index in the project."
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(TileSetDefs) To UBound(TileSetDefs)
            If UCase$(TileSetDefs(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(TileSetDefs) Then
        Err.Raise vbObjectError, , "Tileset " & CStr(Key) & " not found"
    End If
    
    Set TileSetDef = TileSetDefs(I)
End Property

Public Property Get TileSetDefCount() As Integer
Attribute TileSetDefCount.VB_Description = "Return the number of TileSetDef objects in the project."
    If IsEmpty(TileSetDefs) Then
        TileSetDefCount = 0
    Else
        TileSetDefCount = UBound(TileSetDefs) - LBound(TileSetDefs) + 1
    End If
End Property

Public Sub RemoveTileSet(ByVal Key As Variant)
Attribute RemoveTileSet.VB_Description = "Remove a TileSetDef given its name or ordinal position in the project."
    Dim I As Integer
    
    If VarType(Key) = vbString Then
        For I = LBound(TileSetDefs) To UBound(TileSetDefs)
            If UCase$(TileSetDefs(I).Name) = UCase$(Key) Then
                Exit For
            End If
        Next I
    Else
        I = Key
    End If
    
    If I > UBound(TileSetDefs) Then
        Err.Raise vbObjectError, , "Tileset " & CStr(Key) & " not found"
    End If
    
    Do While I < UBound(TileSetDefs)
        Set TileSetDefs(I) = TileSetDefs(I + 1)
        I = I + 1
    Loop
    
    Set TileSetDefs(I) = Nothing
    
    If I > LBound(TileSetDefs) Then
        ReDim Preserve TileSetDefs(LBound(TileSetDefs) To UBound(TileSetDefs) - 1)
    Else
        TileSetDefs = Empty
    End If
    IsDirty = True
    
End Sub

Public Property Get MapCount() As Integer
Attribute MapCount.VB_Description = "Return the number of Map objects in the project."
    If IsEmpty(MapColl) Then
        MapCount = 0
    Else
        MapCount = UBound(MapColl) - LBound(MapColl) + 1
    End If
End Property

Public Property Get MatchDefCount() As Integer
Attribute MatchDefCount.VB_Description = "Return the number of MatchDef objects in the project."
    If IsEmpty(MatchColl) Then
        MatchDefCount = 0
    Else
        MatchDefCount = UBound(MatchColl) - LBound(MatchColl) + 1
    End If
End Property

Public Property Get AnimDefCount() As Integer
Attribute AnimDefCount.VB_Description = "Return the number of AnimDefs in the project"
    If IsEmpty(AnimColl) Then
        AnimDefCount = 0
    Else
        AnimDefCount = UBound(AnimColl) - LBound(AnimColl) + 1
    End If
End Property

Public Sub Save(ByVal FileName As String)
Attribute Save.VB_Description = "Save the project to a plain text file, and all dependent maps and tileset images to their respective files."
    Dim FF As Integer
    Dim I As Integer
    
    FF = FreeFile
    On Error GoTo ProjSaveErr
    If ProjectPath <> "" Then
        ChDir PathFromFile(ProjectPath)
        ChDrive Left$(ProjectPath, 1)
    End If
    For I = 0 To TileSetDefCount - 1
        With TileSetDef(I)
            If Len(.ImagePath) = 0 Then
                MsgBox "Tileset """ & .Name & """ does not have a filename.  Please go to tilesets and save it before saving the project."
                Exit Sub
            End If
            .ImagePath = GetRelativePath(FileName, GetEntirePath(.ImagePath))
            ChDir PathFromFile(FileName)
            ChDrive Left$(FileName, 1)
            If .IsLoaded Then
                .Save
            End If
            If ProjectPath <> "" Then
                ChDir PathFromFile(ProjectPath)
                ChDrive Left$(ProjectPath, 1)
            End If
        End With
    Next
    For I = 0 To MapCount - 1
        If Len(Maps(I).Path) = 0 Then
            MsgBox "Map """ & Maps(I).Name & """ does not have a filename.  Please go to maps and save it before saving the project."
            Exit Sub
        End If
        Maps(I).Path = GetRelativePath(FileName, GetEntirePath(Maps(I).Path))
        ChDir PathFromFile(FileName)
        ChDrive Left$(FileName, 1)
        Maps(I).Save Maps(I).Path
        If ProjectPath <> "" Then
            ChDir PathFromFile(ProjectPath)
            ChDrive Left$(ProjectPath, 1)
        End If
    Next I
    Open FileName For Output As #FF
    Print #FF, "FileVersion=1"
    Print #FF, "TileSetCount=" & CStr(TileSetDefCount)
    For I = 0 To TileSetDefCount - 1
        Print #FF, "TileSetDef=" & CStr(I)
        Print #FF, TileSetDef(I).Serialize
    Next I
    Print #FF, "MapCount=" & CStr(MapCount)
    For I = 0 To MapCount - 1
        Print #FF, "Map=" & CStr(I)
        Print #FF, Maps(I).Serialize
    Next I
    Print #FF, "MatchDefCount=" & CStr(MatchDefCount)
    For I = 0 To MatchDefCount - 1
        Print #FF, "MatchDef=" & CStr(I)
        Print #FF, MatchDefs(I).Serialize
    Next
    Print #FF, "AnimDefCount=" & CStr(AnimDefCount)
    For I = 0 To AnimDefCount - 1
        Print #FF, "AnimDef=" & CStr(I)
        Print #FF, AnimDefs(I).Serialize
    Next
    Print #FF, "CategoryCount=" & CStr(GroupCount)
    For I = 0 To GroupCount - 1
        Print #FF, "CategoryDef=" & CStr(I)
        Print #FF, GroupByIndex(I).Serialize
    Next
    Print #FF, "SolidDefCount=" & CStr(SolidDefCount)
    For I = 0 To SolidDefCount - 1
        Print #FF, "SolidDef=" & CStr(I)
        Print #FF, SolidDefsByIndex(I).Serialize
    Next
    Print #FF, GamePlayer.Serialize
    Print #FF, MediaMgr.Serialize
    Close #FF
    ProjectPath = FileName
    ChDir PathFromFile(ProjectPath)
    ChDrive Left$(ProjectPath, 1)
    IsDirty = False
    Exit Sub

ProjSaveErr:
    MsgBox Err.Description
    Close #FF
End Sub

Public Sub Load(ByVal FileName As String)
Attribute Load.VB_Description = "Load a GDP plain text project file and all dependent map files into memory."
    Dim FF As Integer
    Dim Data As String
    Dim Count As Integer
    Dim I As Integer
    Dim J As Integer
    Dim FileVer As Single
    
    On Error GoTo LoadProjectErr
    
    FF = FreeFile
    ChDir PathFromFile(FileName)
    ChDrive Left$(FileName, 1)
    Open FileName For Binary Access Read As #FF
        Data = Space$(LOF(FF))
        Get #FF, , Data
    Close #FF
    GeneralDeserialize Data, "FileVersion", FileVer
    If FileVer >= 2 Then
        MsgBox "Project file version too new", vbExclamation
        If Len(ProjectPath) Then
            ChDir PathFromFile(ProjectPath)
            ChDrive Left$(ProjectPath, 1)
        End If
        Exit Sub
    End If
    GeneralDeserialize Data, "TileSetCount", Count
    If Count > 0 Then
        ReDim TileSetDefs(0 To Count - 1) As TileSetDef
        For I = 0 To Count - 1
            GeneralDeserialize Data, "TileSetDef", J
            Set TileSetDefs(J) = New TileSetDef
            TileSetDefs(J).Deserialize Data
        Next
    Else
        TileSetDefs = Empty
    End If
    
    GeneralDeserialize Data, "MapCount", Count
    If Count > 0 Then
        ReDim MapColl(0 To Count - 1) As Map
        For I = 0 To Count - 1
            GeneralDeserialize Data, "Map", J
            Set MapColl(J) = New Map
            MapColl(J).Deserialize Data
        Next
    Else
        MapColl = Empty
    End If
    
    GeneralDeserialize Data, "MatchDefCount", Count
    If Count > 0 Then
        ReDim MatchColl(0 To Count - 1) As MatchDef
        For I = 0 To Count - 1
            GeneralDeserialize Data, "MatchDef", J
            Set MatchColl(J) = New MatchDef
            MatchColl(J).Deserialize Data
        Next
    Else
        MatchColl = Empty
    End If
    
    GeneralDeserialize Data, "AnimDefCount", Count
    If Count > 0 Then
        ReDim AnimColl(0 To Count - 1) As AnimDef
        For I = 0 To Count - 1
            GeneralDeserialize Data, "AnimDef", J
            Set AnimColl(J) = New AnimDef
            AnimColl(J).Deserialize Data
        Next
    Else
        AnimColl = Empty
    End If
    
    GeneralDeserialize Data, "CategoryCount", Count
    If Count > 0 Then
        ReDim GroupColl(0 To Count - 1) As Category
        For I = 0 To Count - 1
            GeneralDeserialize Data, "CategoryDef", J
            Set GroupColl(J) = New Category
            GroupColl(J).Deserialize Data
        Next
    Else
        GroupColl = Empty
    End If
    
    GeneralDeserialize Data, "SolidDefCount", Count
    If Count > 0 Then
        ReDim SolidColl(0 To Count - 1) As SolidDef
        For I = 0 To Count - 1
            GeneralDeserialize Data, "SolidDef", J
            Set SolidColl(J) = New SolidDef
            SolidColl(J).Deserialize Data
        Next
    Else
        SolidColl = Empty
    End If
    
    GamePlayer.Deserialize Data
    MediaMgr.Deserialize Data

    If MapCount > 0 Then
        For I = 0 To MapCount - 1
            Maps(I).Load Maps(I).Path
        Next
    End If

    ProjectPath = FileName
    IsDirty = False
    Exit Sub
    
LoadProjectErr:
    MsgBox Err.Description, vbExclamation
    Close #FF
End Sub

Public Function GetApplicableTileMatchCount(TSDef As TileSetDef) As Integer
Attribute GetApplicableTileMatchCount.VB_Description = "Return the number of MatchDef objects that apply to the specified tileset."
    Dim I As Integer
    Dim J As Integer
    
    For I = 0 To MatchDefCount - 1
        If MatchDefs(I).TSDef Is TSDef Then
            J = J + 1
        End If
    Next
    
    GetApplicableTileMatchCount = J
    
End Function

Public Function GetApplicableTileMatch(TSDef As TileSetDef, ByVal Index As Integer) As MatchDef
Attribute GetApplicableTileMatch.VB_Description = "Find and return the Nth MatchDef for the specified tileset where N=Index."
    Dim I As Integer
    Dim J As Integer
    
    For I = 0 To MatchDefCount - 1
        If MatchDefs(I).TSDef Is TSDef Then
            If J = Index Then
                Set GetApplicableTileMatch = MatchDefs(I)
                Exit Function
            End If
            J = J + 1
        End If
    Next
    
End Function

Private Sub Class_Initialize()
    Set GamePlayer = New Player
    Set MediaMgr = New MediaManager
    GamePlayer.InvBarMargin = 5
    GamePlayer.ScrollMarginX = 32
    GamePlayer.ScrollMarginY = 32
End Sub

Private Sub Class_Terminate()
    Dim Idx As Integer
    
    For Idx = 0 To MapCount - 1
        Maps(Idx).ReleaseLayerRefs
        Maps(Idx).ReleaseSpriteRefs
        Maps(Idx).ReleaseSpriteTemplateRefs
        Maps(Idx).ReleaseRefs
    Next
End Sub
