VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnimDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Animated tile definition. Indexed and owned by project."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: AnimDef.cls - Tile Animation Definition Class
'
'======================================================================

Option Explicit

Dim Frames As Variant
Dim Delays As Variant
Public BaseTile As Byte
Attribute BaseTile.VB_VarProcData = ";Behavior"
Attribute BaseTile.VB_VarDescription = "The tile value to become animated"
Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name of the AnimDef object as indexed by the project"
Public MapName As String
Attribute MapName.VB_VarProcData = ";Relational"
Attribute MapName.VB_VarDescription = "Name of the map to which the animated tile applies"
Public LayerName As String
Attribute LayerName.VB_VarProcData = ";Relational"
Attribute LayerName.VB_VarDescription = "Name of the layer to which the animation applies"

' Run-time only
Dim CurrentFrame As Integer
Dim CurrentDelay As Integer

Public Function Clone() As AnimDef
Attribute Clone.VB_Description = "Create a copy of the AnimDef"
    Dim Result As New AnimDef
    
    Result.CopyFrom Me
    Set Clone = Result
    
End Function

Public Sub CopyFrom(Def As AnimDef)
Attribute CopyFrom.VB_Description = "Acts as an assignment operator"
    Dim I As Integer
    
    Frames = Empty
    Delays = Empty
    For I = 0 To Def.FrameCount - 1
        InsertFrame I, Def.FrameValue(I), Def.FrameDelay(I)
    Next
    
    BaseTile = Def.BaseTile
    Name = Def.Name
    MapName = Def.MapName
    LayerName = Def.LayerName
    
End Sub

Public Function FrameCount() As Integer
Attribute FrameCount.VB_Description = "The number of animation cells"
    If IsEmpty(Frames) Then
        FrameCount = 0
    Else
        FrameCount = UBound(Frames) - LBound(Frames) + 1
    End If
End Function

Public Property Get FrameValue(ByVal Index As Integer) As Byte
Attribute FrameValue.VB_Description = "Index of the tile to be displayed in a particular animation cell."
Attribute FrameValue.VB_ProcData.VB_Invoke_Property = ";Data"
    FrameValue = Frames(Index)
End Property

Public Property Get FrameDelay(ByVal Index As Integer) As Byte
Attribute FrameDelay.VB_Description = "How many screen updates will a particular cell be displayed."
Attribute FrameDelay.VB_ProcData.VB_Invoke_Property = ";Data"
    FrameDelay = Delays(Index)
End Property

Public Property Let FrameValue(ByVal Index As Integer, ByVal NewVal As Byte)
    Frames(Index) = NewVal
End Property

Public Property Let FrameDelay(ByVal Index As Integer, ByVal NewVal As Byte)
    Delays(Index) = NewVal
End Property

Public Sub InsertFrame(ByVal BeforeIndex As Integer, ByVal Value As Byte, ByVal Delay As Byte)
Attribute InsertFrame.VB_Description = "Insert a new frame (cell) into the animation sequence before BeforeIndex."
    Dim I As Integer
    
    If IsEmpty(Frames) Then
        ReDim Frames(0) As Byte
        ReDim Delays(0) As Byte
    Else
        ReDim Preserve Frames(LBound(Frames) To UBound(Frames) + 1)
        ReDim Preserve Delays(LBound(Delays) To UBound(Delays) + 1)
    End If
    
    For I = UBound(Frames) - 1 To BeforeIndex Step -1
        Frames(I + 1) = Frames(I)
        Delays(I + 1) = Delays(I)
    Next
    
    Frames(BeforeIndex) = Value
    Delays(BeforeIndex) = Delay
End Sub

Public Sub RemoveFrame(ByVal Index As Integer)
Attribute RemoveFrame.VB_Description = "Remove a frame (cell) from the animation sequence, shifting subsequent frames and delays back."
    Dim I As Integer
    
    For I = Index To UBound(Frames) - 1
        Frames(I) = Frames(I + 1)
        Delays(I) = Delays(I + 1)
    Next
    
    If UBound(Frames) > LBound(Frames) Then
        ReDim Preserve Frames(LBound(Frames) To UBound(Frames) - 1)
        ReDim Preserve Delays(LBound(Delays) To UBound(Delays) - 1)
    Else
        Frames = Empty
        Delays = Empty
    End If
    
    Reset
End Sub

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Return the contents of the AnimDef object as text that can be put into the project file."
    Dim Data As String
    Dim I As Integer
    
    Data = Hex$(FrameCount)
    If Len(Data) = 1 Then Data = "0" & Data
    
    For I = LBound(Frames) To UBound(Frames)
        Data = Data & IIf(Frames(I) < 16, "0" & Hex$(Frames(I)), Hex$(Frames(I)))
        Data = Data & IIf(Delays(I) < 16, "0" & Hex$(Delays(I)), Hex$(Delays(I)))
    Next
    
    Serialize = GeneralSerialize("Name", "MapName", "LayerName", "BaseTile", "Frames", Name, MapName, LayerName, BaseTile, Data)
End Function

Public Sub Deserialize(Data As String)
Attribute Deserialize.VB_Description = "Convert AnimDef text from the project file into an AnimDef object"
    Dim I As Integer
    Dim FrameDat As String
    
    GeneralDeserialize Data, "Name", "MapName", "LayerName", "BaseTile", "Frames", Name, MapName, LayerName, BaseTile, FrameDat
    
    I = Val("&H" & Left$(FrameDat, 2))
    ReDim Frames(I - 1) As Byte
    ReDim Delays(I - 1) As Byte
    
    For I = 0 To FrameCount - 1
        Frames(I + LBound(Frames)) = Val("&H" & Mid$(FrameDat, I * 4 + 3, 2))
        Delays(I + LBound(Delays)) = Val("&H" & Mid$(FrameDat, I * 4 + 5, 2))
    Next
End Sub

Public Property Get CurTile() As Byte
Attribute CurTile.VB_Description = "Returns the tile index associated with the current frame"
Attribute CurTile.VB_UserMemId = 0
    CurTile = Frames(CurrentFrame)
End Property

Public Sub Advance()
    CurrentDelay = CurrentDelay + 1
    If CurrentDelay >= Delays(CurrentFrame) Then
        CurrentDelay = 0
        CurrentFrame = CurrentFrame + 1
    End If
    If CurrentFrame >= FrameCount Then CurrentFrame = 0
End Sub

Public Sub Reset()
Attribute Reset.VB_Description = "Reset the current frame to the beginning of the animation."
    CurrentFrame = 0
    CurrentDelay = 0
End Sub
