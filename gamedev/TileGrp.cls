VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Simplified (nameless) component of a Category object for generically describing a collection of tiles.  Persistence depends on owner -- never persistent by itself."
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
' File: TileGrp.cls - Simple Tile Group Class
'
'======================================================================

Option Explicit

Private Bits(0 To 31) As Byte
Private Masks(0 To 7) As Byte

Public Function IsMember(ByVal TileIndex As Integer) As Boolean
Attribute IsMember.VB_Description = "Return True if the specified tile index is a member of the TileGroup."
    IsMember = Bits(TileIndex \ 8) And Masks(TileIndex Mod 8)
End Function

Public Sub SetMember(ByVal TileIndex As Integer)
Attribute SetMember.VB_Description = "Set the specified tile index as a member of this TileGroup."
    Bits(TileIndex \ 8) = Bits(TileIndex \ 8) Or Masks(TileIndex Mod 8)
End Sub

Public Sub ClearMember(ByVal TileIndex As Integer)
Attribute ClearMember.VB_Description = "Remove a particular member from the group."
    Bits(TileIndex \ 8) = Bits(TileIndex \ 8) And Not Masks(TileIndex Mod 8)
End Sub

Public Sub ClearAll()
Attribute ClearAll.VB_Description = "Clear this to an empty TileGroup"
    Dim I As Integer
    
    For I = LBound(Bits) To UBound(Bits)
        Bits(I) = 0
    Next
End Sub

Public Function GetMember(ByVal MemberIndex As Integer) As Integer
Attribute GetMember.VB_Description = "Get tile number MemberIndex from the group."
    Dim I As Integer
    Dim J As Integer
    
    Do
        If IsMember(I) Then
            If J = MemberIndex Then
                GetMember = I
                Exit Function
            End If
            J = J + 1
        End If
        I = I + 1
    Loop Until I > 255
    
    GetMember = -1
    
End Function

Public Function GetArray() As Variant
Attribute GetArray.VB_Description = "Return an array of all tile indexes in the TileGroup."
    Dim I As Integer, J As Integer
    Dim V As Variant
    ReDim V(0 To 255) As Byte
    
    For I = 0 To 255
        If IsMember(I) Then
            V(J) = I
            J = J + 1
        End If
    Next
    
    If J > 0 Then
        ReDim Preserve V(0 To J - 1) As Byte
    Else
        V = Empty
    End If
    
    GetArray = V
    
End Function

Public Function IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Return true if there are no members."
    Dim I As Integer
    
    IsEmpty = True
    
    For I = 0 To 31
        If Bits(I) Then IsEmpty = False
    Next
    
End Function

Public Function Serialize() As String
Attribute Serialize.VB_Description = "Store the TileGroup into a string for the Project file."
    Dim I As Integer
    
    For I = 0 To 31
        If Bits(I) >= 16 Then
            Serialize = Serialize & Hex$(Bits(I))
        Else
            Serialize = Serialize & "0" & Hex$(Bits(I))
        End If
    Next
    
End Function

Public Sub Deserialize(ByRef Data As String)
Attribute Deserialize.VB_Description = "Load a TileGroup from a string as stored in the project file."
    Dim I As Integer
    
    For I = 0 To 31
        Bits(I) = Val("&H" & Mid$(Data, I * 2 + 1, 2))
    Next
End Sub

Public Property Get RawByte(ByVal ByteIndex As Integer) As Byte
Attribute RawByte.VB_Description = "Return a byte (0-31) used to store the bits of the tilegroup."
    RawByte = Bits(ByteIndex)
End Property

Public Sub MergeFrom(TG As TileGroup)
Attribute MergeFrom.VB_Description = "Merge the members of another TileGroup into this TileGroup."
    Dim I As Integer
    
    For I = LBound(Bits) To UBound(Bits)
        Bits(I) = Bits(I) Or TG.RawByte(I)
    Next
End Sub

Private Sub Class_Initialize()
    Dim I As Integer
    
    For I = 0 To 7
        Masks(I) = 2 ^ I
    Next
End Sub
