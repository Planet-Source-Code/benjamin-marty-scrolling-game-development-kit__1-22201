VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollisionDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Definition of what collisions are interesting and how to react.  Indexed and owned by map."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: CollDef.cls - Sprite Collision Definition Class
'
'======================================================================

Option Explicit

Public Enum eCollisionFlags
    COLL_PLATFORM = 1
    COLL_SWAPVEL = 2
    COLL_ASTOP = 4
    COLL_ATERM = 8
    COLL_ANEW = 16
    COLL_BSTOP = 32
    COLL_BTERM = 64
    COLL_BNEW = 128
End Enum

Public Enum eCollInvFlags
    COLL_INV_REQUIRE = 1
    COLL_INV_REMOVE = 2
End Enum

Dim TestClasses As Byte
Public Flags As Byte
Attribute Flags.VB_VarUserMemId = 0
Attribute Flags.VB_VarProcData = ";Behavior"
Attribute Flags.VB_VarDescription = "How to react to the collision."
Public SpecialFunction As String
Attribute SpecialFunction.VB_VarProcData = ";Behavior"
Attribute SpecialFunction.VB_VarDescription = "Indicates a special function to execute when the collision occurs."
Public InvFlags As Byte
Attribute InvFlags.VB_VarDescription = "Indicates how the collision relates to inventory"
Public InvItem As Byte
Attribute InvItem.VB_VarDescription = "Which inventory item is related to the collision"
Public InvUseCount As Integer
Attribute InvUseCount.VB_VarDescription = "How many inventory items are required to alter the collisions effect."
Public Media As String
Attribute Media.VB_VarDescription = "Name of media clip to play on collision"

Public Property Get ClassA() As Integer
Attribute ClassA.VB_Description = "Number 0 to 15 indicating the collision class of the first part."
Attribute ClassA.VB_ProcData.VB_Invoke_Property = ";Data"
    ClassA = Int(TestClasses / 16)
End Property

Public Property Let ClassA(ByVal Index As Integer)
    TestClasses = (TestClasses And 15) Or CByte(Index * 16)
End Property

Public Property Get ClassB() As Integer
Attribute ClassB.VB_Description = "Number 0 to 15 indicating the collision class of the second part."
Attribute ClassB.VB_ProcData.VB_Invoke_Property = ";Data"
    ClassB = TestClasses Mod 16
End Property

Public Property Let ClassB(ByVal Index As Integer)
    TestClasses = (TestClasses And 240) Or CByte(Index)
End Property

Public Sub Save(ByVal FNum As Integer)
Attribute Save.VB_Description = "Writes the CollisionDef to a binary file."
    Put #FNum, , TestClasses
    Put #FNum, , Flags
    SaveString FNum, SpecialFunction
    ' Version 2
    Put #FNum, , InvFlags
    Put #FNum, , InvItem
    Put #FNum, , InvUseCount
    SaveString FNum, Media
End Sub

Public Sub Load(ByVal FNum As Integer)
Attribute Load.VB_Description = "Reads data from a binary file to construct a CollisionDef."
    Get #FNum, , TestClasses
    Get #FNum, , Flags
    SpecialFunction = LoadString(FNum)
    If LoadVersion >= 2 Then
        ' Version 2
        Get #FNum, , InvFlags
        Get #FNum, , InvItem
        Get #FNum, , InvUseCount
        Media = LoadString(FNum)
    End If
End Sub
