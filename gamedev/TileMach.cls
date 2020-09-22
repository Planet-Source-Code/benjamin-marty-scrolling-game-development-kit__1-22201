VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Simplified (nameless) component of a MatchDef object.  Persistent in the project only as a component of a MatchDef."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: TileMach.cls - Tile Match Definition Class
'
'======================================================================

Option Explicit

Private Groups(0 To 14) As New TileGroup

Public Enum MatchGroups
    MG_TOPLEFT = 0
    MG_TOP
    MG_TOPRIGHT
    MG_INNERTOPLEFT
    MG_INNERTOPRIGHT
    MG_LEFT
    MG_CENTER
    MG_RIGHT
    MG_INNERBOTTOMLEFT
    MG_INNERBOTTOMRIGHT
    MG_BOTTOMLEFT
    MG_BOTTOM
    MG_BOTTOMRIGHT
End Enum

Public Function IsMember(ByVal TileIndex As Integer) As Boolean
Attribute IsMember.VB_Description = "Return True if the specified tile index is a member of some group in the TileMatch."
    Dim I As Integer
    
    For I = 0 To 14
        If Groups(I).IsMember(TileIndex) Then
            IsMember = True
            Exit Function
        End If
    Next
    
    IsMember = False
    
End Function

Public Function SetMember(ByVal GroupIndex As Integer, ByVal TileIndex As Integer)
Attribute SetMember.VB_Description = "Set the specified tile index as a member of the specified group."
    Groups(GroupIndex).SetMember TileIndex
End Function

Public Function ClearMember(ByVal GroupIndex As Integer, ByVal TileIndex As Integer)
Attribute ClearMember.VB_Description = "Remove the specified tile from the specified group."
    Groups(GroupIndex).ClearMember TileIndex
End Function

Public Property Get MatchGroup(ByVal GroupIndex As Byte) As TileGroup
Attribute MatchGroup.VB_Description = "Return one of the 15 TileGroup objects from the MatchDef, based on index."
    Set MatchGroup = Groups(GroupIndex)
End Property

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the TileMatch object to a string for the Project file."
    
    Serialize = GeneralSerialize("TLGroup", "TGroup", "TRGroup", "ITLGroup", "ITRGroup", "LGroup", "CGroup", "RGroup", "IBLGroup", "IBRGroup", "BLGroup", "BGroup", "BRGroup", "IDRGroup", "IDLGroup", Groups(0).Serialize, Groups(1).Serialize, Groups(2).Serialize, Groups(3).Serialize, _
                                 Groups(4).Serialize, Groups(5).Serialize, Groups(6).Serialize, Groups(7).Serialize, Groups(8).Serialize, Groups(9).Serialize, Groups(10).Serialize, Groups(11).Serialize, Groups(12).Serialize, Groups(13).Serialize, Groups(14).Serialize)

End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a TileMatch object from a string."
    Dim S(0 To 14) As String
    Dim I As Integer
    
    GeneralDeserialize Data, "TLGroup", "TGroup", "TRGroup", "ITLGroup", "ITRGroup", "LGroup", "CGroup", "RGroup", "IBLGroup", "IBRGroup", "BLGroup", "BGroup", "BRGroup", "IDRGroup", "IDLGroup", S(0), S(1), S(2), S(3), S(4), S(5), S(6), S(7), S(8), S(9), S(10), S(11), S(12), S(13), S(14)
    
    For I = 0 To 14
        Groups(I).Deserialize S(I)
    Next I
    
End Sub
