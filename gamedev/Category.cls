VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Tile category.  Indexed by Tileset and owned by the project."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Category.cls - Named Tile Category Class
'
'======================================================================

Option Explicit

Public TSName As String
Attribute TSName.VB_VarProcData = ";Relational"
Attribute TSName.VB_VarDescription = "Name of the tileset whose tiles are being categorized."
Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name for this category as indexed by the project for a particular tileset."
Public Group As New TileGroup
Attribute Group.VB_VarUserMemId = 0
Attribute Group.VB_VarProcData = ";Data"
Attribute Group.VB_VarDescription = "The group of tiles indexes that make up this named category of tileset tiles."

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Return the contents of the Category as a string for the project file."
    Serialize = GeneralSerialize("GroupName", "TilesetName", "Group", Name, TSName, Group.Serialize)
End Function

Public Sub Deserialize(Data As String)
Attribute Deserialize.VB_Description = "Convert text from the project file into a Category object."
    Dim GrpStr As String
    
    GeneralDeserialize Data, "GroupName", "TilesetName", "Group", Name, TSName, GrpStr
    Group.Deserialize GrpStr
    
End Sub

Public Function GetIndex() As Integer
Attribute GetIndex.VB_Description = "Find the ordinal position of this Category as indexed by Project.GroupByIndex."
    Dim I As Integer
    
    For I = 0 To Prj.GroupCount - 1
        If Prj.GroupByIndex(I) Is Me Then
            GetIndex = I
            Exit Function
        End If
    Next
    GetIndex = -1
End Function

Public Function GetIndexByTileset(ByVal TSName As String) As Integer
Attribute GetIndexByTileset.VB_Description = "Find the ordinal position of this category as indexed by Project.TilesetGroupByIndex."
    Dim I As Integer
   
    For I = 0 To Prj.GroupByTilesetCount(TSName) - 1
        If Prj.TilesetGroupByIndex(TSName, I) Is Me Then
            GetIndexByTileset = I
            Exit Function
        End If
    Next
End Function
