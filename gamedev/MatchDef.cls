VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MatchDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Fully describes a ""TileMatch"" which automatically fits tiles together in the map editor.  Indexed and owned by the project (even though sometimes appears to be indexed by tileset, and should be)."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: MatchDef.cls - Tile Matching Definition Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this MatchDef object as indexed by the project."
Private TM As New TileMatch ' The tiles classified into categories
Private TotalGroup As New TileGroup ' All tiles in the group, including unclassified tiles
Public TSDef As TileSetDef
Attribute TSDef.VB_VarProcData = ";Relational"
Attribute TSDef.VB_VarDescription = "Indicates the tileset to which this MatchDef applies."

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the contents of this MatchDef to a string."
    
    Serialize = GeneralSerialize("Name", "MatchGroup", "TileSet", Name, TotalGroup.Serialize, TSDef.Name) & TM.Serialize

End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a MatchDef object from a string (GDP file)."
    Dim TotalData As String
    Dim TSName As String
    
    GeneralDeserialize Data, "Name", "MatchGroup", "TileSet", Name, TotalData, TSName
    Set TSDef = Prj.TileSetDef(TSName)
    TotalGroup.Deserialize TotalData
    TM.Deserialize Data

End Sub

Public Function MatchMiddleGrp(ByVal TLTile As Byte, ByVal TTile As Byte, ByVal TRTile As Byte, ByVal LTile As Byte, ByVal RTile As Byte, ByVal BLTile As Byte, ByVal BTile As Byte, ByVal BRTile As Byte) As Byte
Attribute MatchMiddleGrp.VB_Description = "Given all tiles around a tile, return the index of the group of tiles that fits in the middle."
    
    MatchMiddleGrp = 255
    
    If AllTiles.IsMember(TTile) Then
        If AllTiles.IsMember(LTile) Then
            If AllTiles.IsMember(RTile) Then
                If AllTiles.IsMember(BTile) Then
                    If AllTiles.IsMember(TLTile) Then
                        If AllTiles.IsMember(TRTile) Then
                            If AllTiles.IsMember(BLTile) Then
                                If AllTiles.IsMember(BRTile) Then
                                    MatchMiddleGrp = 6
                                Else
                                    MatchMiddleGrp = 9
                                End If
                            Else
                                MatchMiddleGrp = 8
                            End If
                        Else
                            If AllTiles.IsMember(BLTile) Then
                                MatchMiddleGrp = 4
                            Else
                                MatchMiddleGrp = 13
                            End If
                        End If
                    Else
                        If AllTiles.IsMember(BRTile) Then
                            MatchMiddleGrp = 3
                        Else
                            MatchMiddleGrp = 14
                        End If
                    End If
                Else
                    If AllTiles.IsMember(TLTile) Then
                        If AllTiles.IsMember(TRTile) Then
                            MatchMiddleGrp = 11
                        Else
                            MatchMiddleGrp = 12
                        End If
                    Else
                        MatchMiddleGrp = 10
                    End If
                End If
            Else
                If AllTiles.IsMember(BTile) Then
                    If AllTiles.IsMember(TLTile) Then
                        If AllTiles.IsMember(BLTile) Then
                            MatchMiddleGrp = 7
                        Else
                            MatchMiddleGrp = 12
                        End If
                    Else
                        MatchMiddleGrp = 2
                    End If
                Else
                    MatchMiddleGrp = 12
                End If
            End If
        Else
            If AllTiles.IsMember(RTile) Then
                If AllTiles.IsMember(BTile) Then
                    If AllTiles.IsMember(TRTile) Then
                        If AllTiles.IsMember(BRTile) Then
                            MatchMiddleGrp = 5
                        Else
                            MatchMiddleGrp = 10
                        End If
                    Else
                        MatchMiddleGrp = 0
                    End If
                Else
                    MatchMiddleGrp = 10
                End If
            End If
        End If
    Else
        If AllTiles.IsMember(LTile) Then
            If AllTiles.IsMember(RTile) Then
                If AllTiles.IsMember(BTile) Then
                    If AllTiles.IsMember(BLTile) Then
                        If AllTiles.IsMember(BRTile) Then
                            MatchMiddleGrp = 1
                        Else
                            MatchMiddleGrp = 2
                        End If
                    Else
                        MatchMiddleGrp = 0
                    End If
                End If
            Else
                MatchMiddleGrp = 2
            End If
        Else
            MatchMiddleGrp = 0
        End If
    End If
    
    If MatchMiddleGrp = 255 Then MatchMiddleGrp = 6

End Function

Public Function MatchMiddleTile(ByVal TLTile As Byte, ByVal TTile As Byte, ByVal TRTile As Byte, ByVal LTile As Byte, ByVal RTile As Byte, ByVal BLTile As Byte, ByVal BTile As Byte, ByVal BRTile As Byte) As Byte
Attribute MatchMiddleTile.VB_Description = "Given all tiles around a tile, return the index of a tile that fits in the middle."
    Dim G As Byte
    Dim V As Variant
    
    G = MatchMiddleGrp(TLTile, TTile, TRTile, LTile, RTile, BLTile, BTile, BRTile)
    V = TM.MatchGroup(G).GetArray
    If IsEmpty(V) Then
        G = 6
        V = TM.MatchGroup(G).GetArray
    End If
    MatchMiddleTile = V(Int(Rnd(1) * (UBound(V) - LBound(V) + 1)) + LBound(V))
    
End Function

Public Property Get TileMatches() As TileMatch
Attribute TileMatches.VB_Description = "Returns a TileMatch that contains the 15 groups of tiles used in the MatchDef."
Attribute TileMatches.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute TileMatches.VB_UserMemId = 0
    Set TileMatches = TM
End Property

Public Property Get AllTiles() As TileGroup
Attribute AllTiles.VB_Description = "Return the TileGroup containing all tiles in the MatchDef"
    Set AllTiles = TotalGroup
End Property

Public Sub UpdateTotalGroup()
Attribute UpdateTotalGroup.VB_Description = "Recompose the group returned by AllTiles by ORing all groups in TileMatches into a single group."
    Dim I As Integer
    
    For I = 0 To 14
        TotalGroup.MergeFrom TM.MatchGroup(I)
    Next
    
End Sub
