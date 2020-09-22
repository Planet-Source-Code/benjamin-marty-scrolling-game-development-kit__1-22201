VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The only class directly externally creatable by COM; intended as the main access point for COM.  Ensures created objects are from same instance."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Engine.cls - GameDev Publically Creatable Engine Class
'
'======================================================================

Option Explicit

Public Property Get ProjectObj() As GameProject
Attribute ProjectObj.VB_Description = "Return the GameDev global project object."
    Set ProjectObj = Prj
End Property

Public Function NewAnimDef() As AnimDef
Attribute NewAnimDef.VB_Description = "Create a new AnimDef in the current instance of GameDev and return it."
    Set NewAnimDef = New AnimDef
End Function

Public Function NewCollisionDef() As CollisionDef
Attribute NewCollisionDef.VB_Description = "Create a new CollisionDef in the current instance of GameDev and return it."
    Set NewCollisionDef = New CollisionDef
End Function

Public Function NewMap() As Map
Attribute NewMap.VB_Description = "Create a new Map in the current instance of GameDev and return it."
    Set NewMap = New Map
End Function

Public Function NewMapEdit() As MapEdit
Attribute NewMapEdit.VB_Description = "Create a new MapEdit in the current instance of GameDev and return it."
    Set NewMapEdit = New MapEdit
End Function

Public Function NewMapInteract() As MapInteract
Attribute NewMapInteract.VB_Description = "Create a new MapInteract in the current instance of GameDev and return it."
    Set NewMapInteract = New MapInteract
End Function

Public Function NewPath() As Path
Attribute NewPath.VB_Description = "Create a new Path in the current instance of GameDev and return it."
    Set NewPath = New Path
End Function

Public Property Get HostObj() As ScriptHost
Attribute HostObj.VB_Description = "Return the GameDev global scripting host object."
    Set HostObj = GameHost
End Property

Public Function NewSpecialFunction() As SpecialFunction
Attribute NewSpecialFunction.VB_Description = "Create a new SpecialFunction in the current instance of GameDev and return it."
    Set NewSpecialFunction = New SpecialFunction
End Function

Public Function NewSprite() As Sprite
Attribute NewSprite.VB_Description = "Create a new Sprite in the current instance of GameDev and return it."
    Set NewSprite = New Sprite
End Function

Public Function NewSpriteDef() As SpriteDef
Attribute NewSpriteDef.VB_Description = "Create a new SpriteDef in the current instance of GameDev and return it."
    Set NewSpriteDef = New SpriteDef
End Function

Public Function NewSpriteTemplate() As SpriteTemplate
Attribute NewSpriteTemplate.VB_Description = "Create a new SpriteTemplate in the current instance of GameDev and return it."
    Set NewSpriteTemplate = New SpriteTemplate
End Function

Public Function NewTileEdit() As TileEdit
Attribute NewTileEdit.VB_Description = "Create a new TileEdit in the current instance of GameDev and return it."
    Set NewTileEdit = New TileEdit
End Function

Public Function NewTileGroup() As TileGroup
Attribute NewTileGroup.VB_Description = "Create a new TileGroup in the current instance of GameDev and return it."
    Set NewTileGroup = New TileGroup
End Function

Public Function NewTileMatch()
Attribute NewTileMatch.VB_Description = "Create a new TileMatch in the current instance of GameDev and return it."
    Set NewTileMatch = New TileMatch
End Function

Public Property Get GameDevForm(ByVal Name As String) As Object
Attribute GameDevForm.VB_Description = "Returns a form object from GameDev's user interface."
    Select Case Name
    Case "frmAbout"
        Set GameDevForm = frmAbout
    Case "frmCollisons"
        Set GameDevForm = frmCollisions
    Case "frmCtrlConfig"
        Set GameDevForm = frmCtrlConfig
    Case "frmGroupTiles"
        Set GameDevForm = frmGroupTiles
    Case "frmManageMedia"
        Set GameDevForm = frmManageMedia
    Case "frmMapEdit"
        Set GameDevForm = frmMapEdit
    Case "frmMatchTile"
        Set GameDevForm = frmMatchTile
    Case "frmPlayer"
        Set GameDevForm = frmPlayer
    Case "frmProject"
        Set GameDevForm = frmProject
    Case "frmSplash"
        Set GameDevForm = frmSplash
    Case "frmSprites"
        Set GameDevForm = frmSprites
    Case "frmTileAnim"
        Set GameDevForm = frmTileAnim
    Case "frmTileImport"
        Set GameDevForm = frmTileImport
    Case "frmTSDisplay"
        Set GameDevForm = frmTSDisplay
    Case "frmTSEdit"
        Set GameDevForm = frmTSEdit
    Case "frmTSMatching"
        Set GameDevForm = frmTSMatching
    Case "frmShortcut"
        Set GameDevForm = frmShortcut
    Case "frmOptions"
        Set GameDevForm = frmOptions
    End Select
End Property

Public Property Get CurrentDisplay() As Object
Attribute CurrentDisplay.VB_Description = "Returns the global reference to the current full screen display object if it is open."
    Set CurrentDisplay = CurDisp
End Property

Public Function NewGameDevObj(ByVal ObjType As String) As Object
Attribute NewGameDevObj.VB_Description = "Create a new object in the current instance of GameDev, given the type name as a string, and return it."
    Select Case ObjType
    Case "AnimDef"
        Set NewGameDevObj = New AnimDef
    Case "CollisionDef"
        Set NewGameDevObj = New CollisionDef
    Case "Map"
        Set NewGameDevObj = New Map
    Case "MapEdit"
        Set NewGameDevObj = New MapEdit
    Case "MapInteract"
        Set NewGameDevObj = New MapInteract
    Case "Path"
        Set NewGameDevObj = New Path
    Case "SpecialFunction"
        Set NewGameDevObj = New SpecialFunction
    Case "Sprite"
        Set NewGameDevObj = New Sprite
    Case "SpriteDef"
        Set NewGameDevObj = New SpriteDef
    Case "SpriteTemplate"
        Set NewGameDevObj = New SpriteTemplate
    Case "TileEdit"
        Set NewGameDevObj = New TileEdit
    Case "TileGroup"
        Set NewGameDevObj = New TileGroup
    Case "TileMatch"
        Set NewGameDevObj = New TileMatch
    Case "NewDisplay"
        Set NewGameDevObj = New BMDXDisplay
    End Select
End Function

Public Function GetGameDevObj(ByVal ObjName As String) As Object
Attribute GetGameDevObj.VB_Description = "Return an object based on the name given as a string."
    Select Case ObjName
    Case "ProjectObj"
        Set GetGameDevObj = Prj
    Case "CurrentDisplay"
        Set GetGameDevObj = CurDisp
    Case "HostObj"
        Set GetGameDevObj = GameHost
    End Select
End Function
