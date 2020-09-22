VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SolidDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Named set of Category objects classified into different shapes defining the shape of solid tiles.  Indexed by tileset, owned by project."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: SolidDef.cls - Solidity Definition Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of the SolidDef as indexed by the project."
Public TSName As String
Attribute TSName.VB_VarProcData = ";Relational"
Attribute TSName.VB_VarDescription = "Name of the tileset to which this SoidDef applies."
Public Solid As Category
Attribute Solid.VB_VarProcData = ";Data"
Attribute Solid.VB_VarDescription = "Category object containing completely solid tiles."
Public Uphill As Category
Attribute Uphill.VB_VarProcData = ";Data"
Attribute Uphill.VB_VarDescription = "Category object containing tiles sloped up and right, solid below."
Public Downhill As Category
Attribute Downhill.VB_VarProcData = ";Data"
Attribute Downhill.VB_VarDescription = "Category object containing tiles sloped down and right, solid below."
Public UpCeil As Category
Attribute UpCeil.VB_VarProcData = ";Data"
Attribute UpCeil.VB_VarDescription = "Category object containing tiles sloped up and right, solid above."
Public DownCeil As Category
Attribute DownCeil.VB_VarProcData = ";Data"
Attribute DownCeil.VB_VarDescription = "Category object containing tiles sloped down and right, solid above."

Public Sub Remove()
Attribute Remove.VB_Description = "Remove the SolidDef object from the project."
    Prj.RemoveSolidDef Me
End Sub

Public Function GetIndexByTileset(ByVal TSName As String) As Integer
Attribute GetIndexByTileset.VB_Description = "Find the index of the SolidDef as indexed by Project.SolidDefIndexByTileset."
    Dim I As Integer
    Dim Cnt As Integer
    
    For I = 0 To Prj.SolidDefCount - 1
        If Prj.SolidDefsByIndex(I).TSName = TSName Then
            If Prj.SolidDefsByIndex(I) Is Me Then
                GetIndexByTileset = Cnt
                Exit Function
            End If
            Cnt = Cnt + 1
        End If
    Next I
    
End Function

Public Function GetIndex() As Integer
Attribute GetIndex.VB_Description = "Find the index of the SolidDef as indexed by the project."
    Dim I As Integer
    
    For I = 0 To Prj.SolidDefCount - 1
        If Prj.SolidDefsByIndex(I) Is Me Then
            GetIndex = I
            Exit Function
        End If
    Next
End Function

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Return the contents of the SolidDef object as a string for the project file."
    Dim UCName As String
    Dim SName As String
    Dim UName As String
    Dim DName As String
    Dim DCName As String
    
    If Not Solid Is Nothing Then
        SName = Solid.Name
    End If
    If Not Uphill Is Nothing Then
        UName = Uphill.Name
    End If
    If Not Downhill Is Nothing Then
        DName = Downhill.Name
    End If
    If Not UpCeil Is Nothing Then
        UCName = UpCeil.Name
    End If
    If Not DownCeil Is Nothing Then
        DCName = DownCeil.Name
    End If
    Serialize = GeneralSerialize("SolidDefName", "TileSetName", "Solid", "Uphill", "DownHill", "UpCeil", "DownCeil", Name, TSName, SName, UName, DName, UCName, DCName)
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a SolidDef from a project file string."
    Dim sl As String, uh As String, dh As String, uc As String, dc As String
    GeneralDeserialize Data, "SolidDefName", "TileSetName", "Solid", "Uphill", "DownHill", "UpCeil", "DownCeil", Name, TSName, sl, uh, dh, uc, dc
    If Prj.GroupExists(sl, TSName) Then Set Solid = Prj.Groups(sl, TSName) Else Set Solid = Nothing
    If Prj.GroupExists(uh, TSName) Then Set Uphill = Prj.Groups(uh, TSName) Else Set Uphill = Nothing
    If Prj.GroupExists(dh, TSName) Then Set Downhill = Prj.Groups(dh, TSName) Else Set Downhill = Nothing
    If Prj.GroupExists(uc, TSName) Then Set UpCeil = Prj.Groups(uc, TSName) Else Set UpCeil = Nothing
    If Prj.GroupExists(dc, TSName) Then Set DownCeil = Prj.Groups(dc, TSName) Else Set DownCeil = Nothing
End Sub
