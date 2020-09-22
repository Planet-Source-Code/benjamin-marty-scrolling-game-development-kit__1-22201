VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A series of points on a layer that a sprite can follow.  Indexed and owned by map."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: Path.cls - Sprite Path Definition Class
'
'======================================================================

Option Explicit

Public Name As String
Attribute Name.VB_VarUserMemId = -518
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "The name of this path as indexed by the Map."
Public LayerName As String
Attribute LayerName.VB_VarProcData = ";Relational"
Attribute LayerName.VB_VarDescription = "Return the name of the layer to which this path applies."
Private Points As Variant

Public Property Get PointCount() As Integer
Attribute PointCount.VB_Description = "Return the number of points in the Path."
    If IsEmpty(Points) Then
        PointCount = 0
    Else
        PointCount = UBound(Points, 2) - LBound(Points, 2) + 1
    End If
End Property

Public Sub AddPoint(ByVal X As Long, ByVal Y As Long)
Attribute AddPoint.VB_Description = "Add an endpoint to this Path"
    Dim PC As Integer
    
    If PointCount = 0 Then
        ReDim Points(0 To 1, 0 To 0) As Long
        PC = 0
    Else
        PC = PointCount
        ReDim Preserve Points(1, PC) As Long
    End If
    Points(0, PC) = X
    Points(1, PC) = Y
End Sub

Public Property Get PointX(ByVal Index As Long) As Long
Attribute PointX.VB_Description = "Return the X coordnate of the point specified by index."
    PointX = Points(0, Index)
End Property

Public Property Get PointY(ByVal Index As Long) As Long
Attribute PointY.VB_Description = "Return the Y coordnate of the point specified by index."
    PointY = Points(1, Index)
End Property

Public Function GetUsedBy() As SpriteDef
Attribute GetUsedBy.VB_Description = "Find and return the first SpriteDef that uses this path."
    Dim I As Integer
    Dim J As Integer
    
    For I = 0 To Prj.MapCount - 1
        For J = 0 To Prj.Maps(I).SpriteDefCount - 1
            If Prj.Maps(I).SpriteDefs(J).rPath Is Me Then
                Set GetUsedBy = Prj.Maps(I).SpriteDefs(J)
                Exit Function
            End If
        Next J
    Next I
    
End Function

Public Sub RemovePoint(ByVal Index As Integer)
Attribute RemovePoint.VB_Description = "Remove a point from this path, specified by index, shifting the rest back."
    Dim I As Integer
    
    If PointCount > 1 Then
        For I = Index To PointCount - 2
            Points(0, I) = Points(0, I + 1)
            Points(1, I) = Points(1, I + 1)
        Next
        ReDim Preserve Points(1, PointCount - 2)
    Else
        Points = Empty
    End If
End Sub

Public Sub OffsetBy(ByVal X As Long, ByVal Y As Long)
Attribute OffsetBy.VB_Description = "Move all points in the path by the specified offset."
   Dim I As Integer
   
   For I = 0 To PointCount - 1
      Points(0, I) = Points(0, I) + X
      Points(1, I) = Points(1, I) + Y
   Next
   
End Sub
