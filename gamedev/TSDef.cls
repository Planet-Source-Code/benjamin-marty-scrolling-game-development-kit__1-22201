VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileSetDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes the parameters and filename (bitmap) of a tileset.  Indexed and owned by project."
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
' File: TSDef.cls - Tileset Definition Class
'
'======================================================================

Option Explicit

' Project-File Persistent
Public ImagePath As String
Attribute ImagePath.VB_VarProcData = ";Text"
Attribute ImagePath.VB_VarDescription = "Path to the image file (if saved)."
Public TileWidth As Integer
Attribute TileWidth.VB_VarDescription = "Get/set the width of the tiles in this tileset."
Public TileHeight As Integer
Attribute TileHeight.VB_VarDescription = "Get/set the height of the tiles in this tileset."
Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this TileSetDef as indexed by the project."
Public LoadedInstance As IBMDXTileSet
Attribute LoadedInstance.VB_VarProcData = ";Relational"
Attribute LoadedInstance.VB_VarDescription = "If the graphics are loaded into display memory, return the existing BMDXTileset."

' Not Project-File Persistent
Public Image As StdPicture
Attribute Image.VB_VarDescription = "The picture of this tileset (if IsLoaded)."
Public IsDirty As Boolean
Attribute IsDirty.VB_VarDescription = "True if TileSetDef has not been saved since last change."

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the parameters of this TileSetDef to a string for the project file."
    Serialize = GeneralSerialize("Name", "ImagePath", "TileWidth", "TileHeight", Name, ImagePath, TileWidth, TileHeight)
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a TileSetDef from a string."
    GeneralDeserialize Data, "name", "imagepath", "tilewidth", "tileheight", Name, ImagePath, TileWidth, TileHeight
End Sub

Public Sub Load()
Attribute Load.VB_Description = "Load the graphics for this TileSetDef from ImagePath."
    Set Image = LoadPicture(ImagePath)
    If Image.Type = 0 Then Set Image = Nothing
    IsDirty = False
End Sub

Public Sub Unload()
Attribute Unload.VB_Description = "Unload the graphics for this tileset from memory."
    Set Image = Nothing
End Sub

Public Sub Save()
Attribute Save.VB_Description = "Save the graphics of this TileSetDef to an image file (ImagePath)."
    If Image Is Nothing Then Err.Raise vbObjectError, , "No Image"
    SavePicture Image, ImagePath
    IsDirty = False
End Sub

Public Property Get IsLoaded() As Boolean
Attribute IsLoaded.VB_Description = "True if the graphics for this TileSetDef are loaded."
    IsLoaded = Not (Image Is Nothing)
End Property

Public Function GetIndex() As Integer
Attribute GetIndex.VB_Description = "Find the index of this TileSetDef in the Project."
    Dim I As Integer
    
    For I = 0 To Prj.TileSetDefCount - 1
        If Prj.TileSetDef(I) Is Me Then
            GetIndex = I
            Exit Function
        End If
    Next I
    
    GetIndex = -1
    
End Function
