VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Interaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Indicates how the player interacts with a particular category of map tiles with respect to inventory.  Indexed and owned by map."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Interact.cls - Player-Map Interaction Definition Class
'
'======================================================================

Option Explicit

Public Enum TileReactions
    REACT_NONE
    REACT_ADDTOINVENTORY
    REACT_SUBTRACTFROMINVENTORY
    REACT_USEINVENTORY
End Enum

Public TouchCategory As Category
Attribute TouchCategory.VB_VarDescription = "Which tiles trigger this interaction when touched."
Public Flags As Integer  ' See InteractionFlags in SpecialFunction class
Attribute Flags.VB_VarProcData = ";Behavior"
Attribute Flags.VB_VarDescription = "Contains a set of InteractionFlags that control aspects of the interaction."
Public Reaction As TileReactions
Attribute Reaction.VB_VarProcData = ";Behavior"
Attribute Reaction.VB_VarDescription = "How the inventory interacts with this event."
Public InvItem As Byte
Attribute InvItem.VB_VarDescription = "The inventory item involved in this interaction."
Public ReplaceTile As Byte
Attribute ReplaceTile.VB_VarDescription = "The new tile index that will appear on the map layer after this interaction."
Public Media As String
Attribute Media.VB_VarDescription = "Name of media clip to play on tile interaction"

Public Function Save(ByVal nFileNum As Integer) As String
Attribute Save.VB_Description = "Write the Interaction object to a binary file."
    Dim CategoryName As String
    Dim TSName As String
    
    If TouchCategory Is Nothing Then
        CategoryName = ""
        TSName = ""
    Else
        CategoryName = TouchCategory.Name
        TSName = TouchCategory.TSName
    End If
    SaveString nFileNum, CategoryName
    SaveString nFileNum, TSName
    Put #nFileNum, , Flags
    Put #nFileNum, , Reaction
    Put #nFileNum, , InvItem
    Put #nFileNum, , ReplaceTile
    ' Version 2
    SaveString nFileNum, Media

End Function

Public Sub Load(ByVal nFileNum As Integer)
Attribute Load.VB_Description = "Load an Interaction from a binary file."
    Dim CategoryName As String
    Dim TSName As String
    
    CategoryName = LoadString(nFileNum)
    TSName = LoadString(nFileNum)
    Get #nFileNum, , Flags
    Get #nFileNum, , Reaction
    Get #nFileNum, , InvItem
    Get #nFileNum, , ReplaceTile
    If Len(CategoryName) Then
        Set TouchCategory = Prj.Groups(CategoryName, TSName)
    Else
        Set TouchCategory = Nothing
    End If
    ' Version 2
    If LoadVersion >= 2 Then
        Media = LoadString(nFileNum)
    End If
End Sub
