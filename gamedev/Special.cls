VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpecialFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines the conditions and reaction for a special interaction between the player and a rectangle on the map.  Owned and indexed by map."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Special.cls - Special Function Definition Class
'
'======================================================================

Option Explicit

Public Enum SpecialFuncs
    SPECIAL_MESSAGE = 1
    SPECIAL_SWITCHSPRITE
    SPECIAL_SWITCHMAP
    SPECIAL_TELEPORT
    SPECIAL_ALTERMAP
    SPECIAL_CREATESPRITE
    SPECIAL_EVENT
    SPECIAL_DELETESPRITE
End Enum

Public Enum InteractionFlags
    INTFL_INITIALTOUCH = 1
    INTFL_REMOVEIFACT = 2
    INTFL_REMOVEALWAYS = 4
    INTFL_RAISEEVENT = 8
    INTFL_ACTONUP = 16
    INTFL_ACTONBUTTON = 32
    INTFL_ACTONDOWN = 64
    INTFL_SWAPCONTROL = 128
    INTFL_NEWINSTANCE = 256
    INTFL_DELETEOLD = 512
    INTFL_OLDLOCATION = 1024
    INTFL_OVERRIDEPOSITION = 128
    INTFL_RELATIVETELEPORT = 128
    INTFL_RELATIVETOPLAYER = 256
    INTFL_CHANGEBGMEDIA = 2048
    INTFL_STOPMEDIA = 4096
    INTFL_FUNCREMOVEINV = 8192
    INTFL_ACTONSTARTUP = 16384
    INTFL_DELETEMANY = 128
End Enum

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name of the SpecialFunction as indexed by the Map."
Public LayerIndex As Integer
Attribute LayerIndex.VB_VarProcData = ";Relational"
Attribute LayerIndex.VB_VarDescription = "Index of the layer in this map to which the SpecialFunction applies."
Public TileLeft As Integer
Attribute TileLeft.VB_VarDescription = "The left coordinate of the special function's activation rectangle."
Public TileTop As Integer
Attribute TileTop.VB_VarDescription = "The top coordinate of the special function's activation rectangle."
Public TileRight As Integer
Attribute TileRight.VB_VarDescription = "The right coordinate of the special function's activation rectangle."
Public TileBottom As Integer
Attribute TileBottom.VB_VarDescription = "The bottom coordinate of the special function's activation rectangle."
Public Flags As Integer ' See InteractionFlags above
Attribute Flags.VB_VarProcData = ";Behavior"
Attribute Flags.VB_VarDescription = "Bitmask of InteractionFlags dictating various aspects how the function behaves."
Public FuncType As SpecialFuncs
Attribute FuncType.VB_VarProcData = ";Behavior"
Attribute FuncType.VB_VarDescription = "Type of action this special function takes."
Public Value As String
Attribute Value.VB_VarUserMemId = 0
Attribute Value.VB_VarProcData = ";Text"
Attribute Value.VB_VarDescription = "The name of a target of the SpecialFunction.  Meaning depends on FuncType."
Public SpriteName As String
Attribute SpriteName.VB_VarProcData = ";Data"
Attribute SpriteName.VB_VarDescription = "Name of the sprite relavent to this SpecialFunction (if any)."
Public DestX As Long
Attribute DestX.VB_VarDescription = "Target coordinate of the special function.  Units depends on function."
Public DestY As Long
Attribute DestY.VB_VarDescription = "Target coordinate of the special function.  Units depends on function."
Public MediaName As String
Attribute MediaName.VB_VarProcData = ";Relational"
Attribute MediaName.VB_VarDescription = "Name of the MediaClip (if any) that plays when the function is activated."
Public InvItem As Byte
Attribute InvItem.VB_VarDescription = "Index of the inventory item used by this function."
Public InvUseCount As Integer
Attribute InvUseCount.VB_VarDescription = "How many of InvItem are used to activate the function."

Public Sub Save(ByVal nFileNum As Integer)
Attribute Save.VB_Description = "Write the SpecialFunction to a binary file."
    SaveString nFileNum, Name
    Put #nFileNum, , LayerIndex
    Put #nFileNum, , TileLeft
    Put #nFileNum, , TileTop
    Put #nFileNum, , TileRight
    Put #nFileNum, , TileBottom
    Put #nFileNum, , Flags
    Put #nFileNum, , FuncType
    SaveString nFileNum, Value
    SaveString nFileNum, SpriteName
    Put #nFileNum, , DestX
    Put #nFileNum, , DestY
    SaveString nFileNum, MediaName
    Put #nFileNum, , InvItem
    Put #nFileNum, , InvUseCount
End Sub

Public Sub Load(ByVal nFileNum As Integer)
Attribute Load.VB_Description = "Load a SpecialFunction from the binary map file."
    Name = LoadString(nFileNum)
    Get #nFileNum, , LayerIndex
    Get #nFileNum, , TileLeft
    Get #nFileNum, , TileTop
    Get #nFileNum, , TileRight
    Get #nFileNum, , TileBottom
    Get #nFileNum, , Flags
    Get #nFileNum, , FuncType
    Value = LoadString(nFileNum)
    SpriteName = LoadString(nFileNum)
    Get #nFileNum, , DestX
    Get #nFileNum, , DestY
    MediaName = LoadString(nFileNum)
    Get #nFileNum, , InvItem
    Get #nFileNum, , InvUseCount
End Sub

