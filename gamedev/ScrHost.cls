VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Script hosting object.  One global instance exists.  Adds an Engine object to the script's global namespace."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: ScrHost.cls - Microsoft® ActiveX® Script Hosting Class
'
'======================================================================

Option Explicit

Dim WithEvents HostObj As SCRHOSTLib.Host
Attribute HostObj.VB_VarHelpID = -1
Dim nScriptTimer As Single
Public ScriptTimeOutSeconds As Single
Attribute ScriptTimeOutSeconds.VB_VarDescription = "How many seconds may elapse before GameDev attempts to abort the script with a timeout error."
Dim bInterrupted As Boolean
Dim ScriptText() As String
Public StartScript As Integer
Attribute StartScript.VB_VarDescription = "If set (>=0), clear out the current script and start a new one (mainly to sink new events). "
Public bScriptIsRunning As Boolean
Attribute bScriptIsRunning.VB_VarDescription = "Indicates if script is running."
Private ScriptError As String
Private bInitialUse As Boolean

Public TempStorage As Variant ' For script to store objects/variables accross RunStartStripts

Private Sub Class_Initialize()
    ScriptTimeOutSeconds = 2
    Set HostObj = New SCRHOSTLib.Host
    bInitialUse = True
End Sub

Public Property Get GameDevTypeLibFile() As String
Attribute GameDevTypeLibFile.VB_Description = "Get the file containing the type library for GameDev (Path to GameDev.EXE)."
    GameDevTypeLibFile = TypeLibFromLibID("{D0BC0A98-5AF6-4B55-9DCB-F6ABB4895D9D}", 1, 0)
End Property

Public Sub InitScript(ByVal StartupScript As String)
Attribute InitScript.VB_Description = "Load a script file into memory.  Split it into its parts if it has ""#Split"" in it."
    Dim SplitPos As Long
    Dim ScriptIndex As Integer
    
    SplitPos = InStr(1, StartupScript, vbCrLf & "#Split", vbTextCompare)
    Do While SplitPos > 0
        ReDim Preserve ScriptText(0 To ScriptIndex)
        ScriptText(ScriptIndex) = Left$(StartupScript, SplitPos - 1)
        SplitPos = InStr(SplitPos + 1, StartupScript, vbCrLf, vbTextCompare)
        If SplitPos > 0 Then
            StartupScript = Mid$(StartupScript, SplitPos + 2)
        End If
        SplitPos = InStr(1, StartupScript, vbCrLf & "#Split", vbTextCompare)
        ScriptIndex = ScriptIndex + 1
    Loop
    
    ReDim Preserve ScriptText(0 To ScriptIndex)
    ScriptText(ScriptIndex) = StartupScript
    
End Sub

Public Sub RunStartScript()
Attribute RunStartScript.VB_Description = "Initialize and run the script indicated by StartScript."
    Dim nStart As Integer
    Dim strErr As String
    
    On Error GoTo RunScriptErr
    
    If StartScript >= 0 Then
        nStart = StartScript
        StartScript = -1
        If bScriptIsRunning Then
            If Not (CurDisp Is Nothing) Then
                CurDisp.Close
                Set CurDisp = Nothing
            End If
            MsgBox "Cannot RunStartScript while script is running", vbExclamation
            Exit Sub
        End If
        
        If Not bInitialUse Then Set HostObj = New SCRHOSTLib.Host
        bInitialUse = False
        
        HostObj.InitEngine VBScript
        HostObj.AddObject "Engine", True, True, True, New Engine, HostObj.CLSIDFromProgID("GameDev.Engine"), GameDevTypeLibFile
        HostObj.AddTypeLib "{D0BC0A98-5AF6-4B55-9DCB-F6ABB4895D9D}", 1, 0
        
        HostObj.SetScriptState SS_STARTED
        HostObj.ParseScriptText ScriptText(nStart), vbNullString
        If StartScript = -2 Then Unload frmProject
    End If
    Exit Sub

RunScriptErr:
    strErr = Err.Description
    If Not (CurDisp Is Nothing) Then
        CurDisp.Close
    End If
    MsgBox strErr, vbExclamation
End Sub

Public Sub ConnectEventsNow()
Attribute ConnectEventsNow.VB_Description = "Finalize all the object event sinks and connect them to the script."
    HostObj.SetScriptState SS_CONNECTED
End Sub

Private Sub Class_Terminate()
    Set HostObj = Nothing
End Sub

Private Function HostObj_InterruptPoll() As Boolean
    If Timer - nScriptTimer > ScriptTimeOutSeconds Then
        HostObj_InterruptPoll = True
        bInterrupted = True
    Else
        HostObj_InterruptPoll = False
        bInterrupted = False
    End If
End Function

Private Sub HostObj_OnEnterScript()
    nScriptTimer = Timer
    bInterrupted = False
    bScriptIsRunning = True
End Sub

Private Sub HostObj_OnLeaveScript()
    bScriptIsRunning = False
End Sub

Private Sub HostObj_OnScriptError(ByVal ulLineNumber As Long, ByVal ichCharPosition As Long, ByVal bstrDescription As String)
    Dim I As Integer

    On Error Resume Next
    CurDisp.Close
    On Error GoTo 0

    ScriptError = "Script stopped at line " & ulLineNumber + 1 & " on character " & ichCharPosition + 1 & ":" & vbCrLf & IIf(bInterrupted, "Script timeout encountered", IIf(Len(bstrDescription), bstrDescription, "No error message available"))
    bInterrupted = False
    
End Sub

Public Function CheckForError() As Boolean
Attribute CheckForError.VB_Description = "Display a MessageBox if there has been a script error, and clear the error."
    If Len(ScriptError) Then
        MsgBox ScriptError, vbExclamation, "Script Error Condition"
        ScriptError = ""
        Set HostObj = Nothing
        CheckForError = True
    End If
End Function

Public Sub AddObjectWithEvents(ByVal Class As String, ByVal Name As String, ByVal TypeLibraryFile As String)
Attribute AddObjectWithEvents.VB_Description = "Create a new instance of an object and add it to the script's namespace and sink events for it."
    HostObj.AddObject Name, False, True, True, CreateObject(Class), HostObj.CLSIDFromProgID(Class), TypeLibraryFile
End Sub

Public Function TypeLibFromLibID(ByVal LibID As String, ByVal VerMajor As Integer, ByVal VerMinor As Integer) As String
Attribute TypeLibFromLibID.VB_Description = "Given a LibID in the format {00000200-0000-0010-8000-00AA006D2EA4} return the file."
    TypeLibFromLibID = HostObj.TypeLibFromLibID(LibID, VerMajor, VerMinor)
End Function

Public Sub SinkObjectEvents(ObjInstance As Object, ByVal Name As String, Optional ByVal ProgId As String, Optional ByVal TypeLibraryFile As String)
Attribute SinkObjectEvents.VB_Description = "Sink an object's events with the given name.  (Must be before ConnectEventsNow to work.)"
    On Error GoTo SinkErr
    
    If Len(TypeLibraryFile) = 0 Then
        TypeLibraryFile = GameDevTypeLibFile
    End If
        
    If TypeOf ObjInstance Is AnimDef Then
        ProgId = "GameDev.AnimDef"
    ElseIf TypeOf ObjInstance Is Category Then
        ProgId = "GameDev.Category"
    ElseIf TypeOf ObjInstance Is CollisionDef Then
        ProgId = "GameDev.CollisionDef"
    ElseIf TypeOf ObjInstance Is Engine Then
        ProgId = "GameDev.Engine"
    ElseIf TypeOf ObjInstance Is GameProject Then
        ProgId = "GameDev.GameProject"
    ElseIf TypeOf ObjInstance Is Interaction Then
        ProgId = "GameDev.Interaction"
    ElseIf TypeOf ObjInstance Is Layer Then
        ProgId = "GameDev.Layer"
    ElseIf TypeOf ObjInstance Is Map Then
        ProgId = "GameDev.Map"
    ElseIf TypeOf ObjInstance Is MapEdit Then
        ProgId = "GameDev.MapEdit"
    ElseIf TypeOf ObjInstance Is MapInteract Then
        ProgId = "GameDev.MapInteract"
    ElseIf TypeOf ObjInstance Is MatchDef Then
        ProgId = "GameDev.MatchDef"
    ElseIf TypeOf ObjInstance Is MediaClip Then
        ProgId = "GameDev.MediaClip"
    ElseIf TypeOf ObjInstance Is MediaManager Then
        ProgId = "GameDev.MediaManager"
    ElseIf TypeOf ObjInstance Is Path Then
        ProgId = "GameDev.Path"
    ElseIf TypeOf ObjInstance Is Player Then
        ProgId = "GameDev.Player"
    ElseIf TypeOf ObjInstance Is ScriptHost Then
        ProgId = "GameDev.ScriptHost"
    ElseIf TypeOf ObjInstance Is SolidDef Then
        ProgId = "GameDev.SolidDef"
    ElseIf TypeOf ObjInstance Is SpecialFunction Then
        ProgId = "GameDev.SpecialFunction"
    ElseIf TypeOf ObjInstance Is Sprite Then
        ProgId = "GameDev.Sprite"
    ElseIf TypeOf ObjInstance Is SpriteDef Then
        ProgId = "GameDev.SpriteDef"
    ElseIf TypeOf ObjInstance Is SpriteTemplate Then
        ProgId = "GameDev.SpriteTemplate"
    ElseIf TypeOf ObjInstance Is TileEdit Then
        ProgId = "GameDev.TileEdit"
    ElseIf TypeOf ObjInstance Is TileGroup Then
        ProgId = "GameDev.TileGroup"
    ElseIf TypeOf ObjInstance Is TileMatch Then
        ProgId = "GameDev.TileMatch"
    ElseIf TypeOf ObjInstance Is TileSetDef Then
        ProgId = "GameDev.TileSetDef"
    ElseIf TypeOf ObjInstance Is BMDXDisplay Then
        ProgId = "BMDXDisplay.BMDXDisplay.1"
        TypeLibraryFile = TypeLibFromLibID("{15DAEC40-B0F5-11D1-BB79-EDF05EE9F06E}", 1, 0)
    End If
    
    HostObj.AddObject Name, False, True, False, ObjInstance, HostObj.CLSIDFromProgID(ProgId), TypeLibraryFile
    Exit Sub
    
SinkErr:
    MsgBox Err.Description, vbExclamation, "SinkObjectEvents"
End Sub

Public Function AsObject(VarObj As Variant) As Object
    Set AsObject = VarObj
End Function

Public Function DoCommand(CmdInfo As Variant) As Variant
    DoCommand = 0 ' 0 = Not Implemented
End Function

Public Function CapturePic(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As IPictureDisp
    Set CapturePic = CapturePicture(hDC, Left, Top, Width, Height)
End Function

Public Sub PasteTileToPic(ByVal Pic As Variant, ByVal Tile As Variant, ByVal X As Integer, ByVal Y As Integer)
    PasteTileToPicture Pic, Tile, X, Y
End Sub

Public Function ExtractPic(ByVal Pic As Variant, ByVal ExLeft As Long, ByVal ExTop As Long, ByVal ExWidth As Long, ByVal ExHeight As Long, Optional ByVal bHighlight As Boolean = False) As StdPicture
    Set ExtractPic = ExtractTile(Pic, ExLeft, ExTop, ExWidth, ExHeight, bHighlight)
End Function

