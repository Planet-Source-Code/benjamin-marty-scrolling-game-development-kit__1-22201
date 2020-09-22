VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Full screen tileset editor class.  Not persistent or owned."
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
' File: TileEdit.cls - Full Screen Tile Editor User Interface Class
'
'======================================================================

Option Explicit

Private Type Button
    rcFrame As RECT
    Caption As String
End Type

Dim Sq2 As Double '= Sqr(2)

Public WithEvents Disp As BMDXDisplay
Attribute Disp.VB_VarHelpID = -1
Dim Tileset As IBMDXTileSet
Public Magnify As Integer
Attribute Magnify.VB_VarDescription = "The magnification level of the grid (linear grid pixels per image pixel)."
Public GridLeft As Integer
Attribute GridLeft.VB_VarDescription = "The left coordinate of the magnified tile view."
Public GridTop As Integer
Attribute GridTop.VB_VarDescription = "The top coordinate of the magnified tile view."
Public bGrid As Boolean
Attribute bGrid.VB_VarDescription = "True if the grid should be drawn in the magnified view."
Public ImageLeft As Long
Attribute ImageLeft.VB_VarDescription = "The left coordinate of the un-magnified image."
Public ImageTop As Long
Attribute ImageTop.VB_VarDescription = "The top coordinate of the un-magnified image."
Private Tiles As StdPicture
Public TileRows As Integer
Attribute TileRows.VB_VarDescription = "Number of rows in the tileset being edited."
Public MouseX As Integer
Attribute MouseX.VB_VarDescription = "Current mouse X coordinate."
Public MouseY As Integer
Attribute MouseY.VB_VarDescription = "Current mouse Y coordinate."
Dim MouseTS As IBMDXTileSet
Dim ButtonPress As Integer
Dim Buttons(0 To 55) As Button
Dim hdcMouseBack As Long
Dim hbmpMouseBack As Long
Dim hbmpNull As Long
Public hDC As Long
Attribute hDC.VB_VarDescription = "The handle to the device context for the display if it is locked."
Dim bHandleEvents As Boolean
Public PalLeft As Long
Attribute PalLeft.VB_VarDescription = "Left coordinate of the palette display."
Public PalTop As Long
Attribute PalTop.VB_VarDescription = "Top coordinate of the palette display."
Public PalCellWidth As Integer
Attribute PalCellWidth.VB_VarDescription = "Width of a cell in the palette."
Public PalCellHeight As Integer
Attribute PalCellHeight.VB_VarDescription = "Height of a cell in the palette."
Public PalRows As Integer
Attribute PalRows.VB_VarDescription = "Number of rows in the palette display."
Public PalCols As Integer
Attribute PalCols.VB_VarDescription = "Number of columns in the palette."
Dim PalEntries() As OLE_COLOR
Public CurPalEntry As Integer
Attribute CurPalEntry.VB_VarDescription = "The index of the currently selected color in the palette."
Public SecPalEntry As Integer
Attribute SecPalEntry.VB_VarDescription = "Index of the currently selected secondary palette color."
Dim TilePic As StdPicture
Public CurTool As Integer
Attribute CurTool.VB_VarDescription = "Number indicating the currently selected drawing tool/mode."
Dim ButtonFont As IFont
Dim RButtonPick As Boolean
Public ClipLeft As Long
Attribute ClipLeft.VB_VarDescription = "The left coordinate of where the copied image is drawn."
Public ClipTop As Long
Attribute ClipTop.VB_VarDescription = "The top coordinate of where the copied image is drawn."
Public RotX As Long
Attribute RotX.VB_VarDescription = "X coordinate for the rotated image info."
Public RotY As Long
Attribute RotY.VB_VarDescription = "Y coordinate for the rotated image info."
Dim ClipPic As StdPicture
Public RotAngle As Integer
Attribute RotAngle.VB_VarDescription = "Currently selected angle of rotation."
Public AAOn As Boolean
Attribute AAOn.VB_VarDescription = "True if Anti-Aliasing is turned on for rotation."
Dim UndoBuf(0 To 4) As StdPicture
Dim CurUndo As Integer
Public CurCell As Integer
Attribute CurCell.VB_VarDescription = "The index of the currently selected tile in the tileset."
Public TSViewTop As Integer
Public TSViewLeft As Integer
Public TSViewWidth As Integer
Public TSViewHeight As Integer
Dim hdcGridBuf As Long
Dim hbmpGridBuf As Long
Dim hbmpGridNull As Long
Dim bTempImage As Boolean
Public Translucent As Integer
Attribute Translucent.VB_VarDescription = "Current translucency setting (0-100)."

Event OnEditInit()
Attribute OnEditInit.VB_Description = "Fired after the editor is initialized."
Event OnEditComplete()
Attribute OnEditComplete.VB_Description = "Fired when the close button is clicked."
Event OnKeyPress(ByVal KeyAscii As Integer)
Attribute OnKeyPress.VB_Description = "Fired when a key is pressed."
Event OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseMove.VB_Description = "Fired when the mouse moves."
Event OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseDown.VB_Description = "Fired when the mouse is pressed."
Event OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Attribute OnMouseUp.VB_Description = "Fired when the mouse is released."

Public Sub DrawGrid()
Attribute DrawGrid.VB_Description = "Draw the grid if it's enabled."
    Dim X As Integer
    Dim Y As Integer
    Dim bTempDC As Boolean
    Dim rcGrid As RECT
    
    On Error GoTo DGErr

    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
        
    If bGrid Then
        With rcGrid
            .Left = GridLeft
            .Top = GridTop
            .Right = GridLeft + Tileset.TileWidth * Magnify + 1
            .Bottom = GridTop + Tileset.TileHeight * Magnify + 1
        End With
        FillRect hDC, rcGrid, GetStockObject(WHITE_BRUSH)
    End If
    
    For Y = 0 To Tileset.TileHeight - 1
        For X = 0 To Tileset.TileWidth - 1
            DrawMagPixel X, Y
        Next
    Next
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    Exit Sub
    
DGErr:
    Disp.Close
    hDC = 0
    MsgBox Err.Description, vbCritical
End Sub

Public Sub DrawMagPixel(ByVal XOff As Integer, ByVal YOff As Integer)
Attribute DrawMagPixel.VB_Description = "Draw a magnified pixel at the specified image coordinates."
    Dim hBrush As Long
    Dim rcPixel As RECT
    Dim bTempDC As Boolean

    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    hBrush = CreateSolidBrush(GetPixel(hDC, ImageLeft + XOff, ImageTop + YOff))
    If bGrid Then
        rcPixel.Left = GridLeft + Magnify * XOff + 1
        rcPixel.Top = GridTop + Magnify * YOff + 1
        rcPixel.Right = rcPixel.Left + Magnify - 1
        rcPixel.Bottom = rcPixel.Top + Magnify - 1
    Else
        rcPixel.Left = GridLeft + Magnify * XOff
        rcPixel.Top = GridTop + Magnify * YOff
        rcPixel.Right = rcPixel.Left + Magnify
        rcPixel.Bottom = rcPixel.Top + Magnify
    End If
    FillRect hDC, rcPixel, hBrush
    DeleteObject hBrush
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
End Sub

Public Sub LoadTile(ByVal Index As Integer)
Attribute LoadTile.VB_Description = "Retrieve a tile from the tileset into the grid/editor."
    Disp.DrawTile Tileset, Index, ImageLeft, ImageTop, 0, 0, 640, 480, False
    hDC = Disp.GetDC
    Set TilePic = CapturePicture(hDC, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight)
    Disp.ReleaseDC
    hDC = 0
End Sub

Public Sub SaveTile(ByVal Index As Integer)
Attribute SaveTile.VB_Description = "Store the tile in the editor to the tileset."
    
    On Error GoTo SaveErr
    With Tileset
        hDC = Disp.GetDC
        Set TilePic = CapturePicture(hDC, ImageLeft, ImageTop, .TileWidth, .TileHeight)
        Disp.ReleaseDC
        hDC = 0
        PasteTileToPicture Tiles, TilePic, (Index Mod .TileColumns) * .TileWidth, (Index \ .TileColumns) * .TileHeight
        .PaintPicture TilePic, (Index Mod .TileColumns) * .TileWidth, (Index \ .TileColumns) * .TileHeight
    End With
    DrawAll
    Exit Sub
    
SaveErr:
    Disp.Close
    MsgBox Err.Description, vbCritical
End Sub

Public Sub Edit(EditTiles As StdPicture, ByVal TileWidth As Integer, ByVal TileHeight As Integer)
Attribute Edit.VB_Description = "Start the tile editor using an existing tileset."
    Dim Fnt As New StdFont
    Dim RowCount As Integer, ColCount As Integer
    
    On Error GoTo EditErr
    
    Fnt.Name = "MS Sans Serif"
    Fnt.Size = 8
    Fnt.Bold = True
    Set ButtonFont = Fnt
    Disp.SetFont Fnt

    RowCount = Disp.HimetricToPixelY(EditTiles.Height) \ TileHeight
    ColCount = Disp.HimetricToPixelX(EditTiles.Width) \ TileWidth
    TileRows = RowCount
    hDC = Disp.GetDC
    MakeBlankTile TileWidth, TileHeight
    Set Tiles = EditTiles
    hdcMouseBack = CreateCompatibleDC(hDC)
    hbmpMouseBack = CreateCompatibleBitmap(hDC, 17, 17)
    hbmpNull = SelectObject(hdcMouseBack, hbmpMouseBack)
    Disp.ReleaseDC
    hDC = 0
    Set Tileset = Disp.CreateTileSet(Tiles, TileWidth, TileHeight)
    Set MouseTS = Disp.CreateTileSet(GetMouseBmp(), 17, 17)
    InitScreen
    RaiseEvent OnEditInit
    Exit Sub
    
EditErr:
    Disp.Close
    MsgBox Err.Description
End Sub

Sub InitScreen()
Attribute InitScreen.VB_Description = "Initialize components of the display that are based on the tileset."
    GridLeft = 0
    GridTop = 0
    
    If Tileset.TileWidth > 32 Or Tileset.TileHeight > 32 Then
        If Tileset.TileWidth > Tileset.TileHeight Then
            Magnify = 256 \ Tileset.TileWidth
        Else
            Magnify = 256 \ Tileset.TileHeight
        End If
        If Magnify < 2 Then Magnify = 2
    Else
        Magnify = 8
    End If
    
    ImageLeft = GridLeft + Tileset.TileWidth * Magnify + Magnify
    ImageTop = 0
    ClipLeft = ImageLeft
    ClipTop = ImageTop + Tileset.TileHeight + 30
    RotX = ImageLeft + Tileset.TileWidth + 30
    RotY = 30
    TSViewLeft = 310
    TSViewTop = 280
    TSViewWidth = 180
    TSViewHeight = Tileset.TileHeight * 3
    CalcButtons
    MakeGridBuf
    bHandleEvents = True
    DrawAll
    SetCursor 0
    PushUndo
End Sub

Public Sub Create(ByVal TileWidth As Integer, ByVal TileHeight As Integer, ByVal RowCount As Integer, ByVal ColCount As Integer)
Attribute Create.VB_Description = "Start the tileset editor creating a new tileset image."
    Dim Fnt As New StdFont
    
    Fnt.Name = "MS Sans Serif"
    Fnt.Size = 8
    Fnt.Bold = True
    Set ButtonFont = Fnt
    Disp.SetFont Fnt
        
    TileRows = RowCount
    hDC = Disp.GetDC
    MakeBlankTile TileWidth, TileHeight
    Set Tiles = CapturePicture(hDC, -TileWidth * ColCount, -TileHeight * RowCount, TileWidth * ColCount, TileHeight * RowCount)
    hdcMouseBack = CreateCompatibleDC(hDC)
    hbmpMouseBack = CreateCompatibleBitmap(hDC, 17, 17)
    hbmpNull = SelectObject(hdcMouseBack, hbmpMouseBack)
    Disp.ReleaseDC
    hDC = 0
    Set Tileset = Disp.CreateTileSet(Tiles, TileWidth, TileHeight)
    Set MouseTS = Disp.CreateTileSet(GetMouseBmp(), 17, 17)
    InitScreen
End Sub

Private Sub Class_Initialize()
    InitPalette
    InitButtons
    
    Sq2 = Sqr(2)
    
    bHandleEvents = False
    bGrid = True
End Sub

Private Sub Class_Terminate()
    Dim I As Integer
    
    If hdcMouseBack Then
        SelectObject hdcMouseBack, hbmpNull
        DeleteObject hbmpMouseBack
        DeleteDC hdcMouseBack
        hdcMouseBack = 0
        hbmpMouseBack = 0
    End If
    DeleteGridBuf
    If hDC Then
        Disp.ReleaseDC
    End If
    I = 0
    If Not (Disp Is Nothing) Then
        Do While I < Disp.TileSets.Count
            If Disp.TileSets.Item(I) Is MouseTS Or Disp.TileSets.Item(I) Is Tileset Then
                Disp.TileSets.Remove I
                I = I - 1
            End If
            I = I + 1
        Loop
    End If
    Set MouseTS = Nothing
    Set Tiles = Nothing
    Set Tileset = Nothing
    Set CurDisp = Nothing
    Set Disp = Nothing
    Set ButtonFont = Nothing
    Exit Sub
End Sub

Private Sub Disp_KeyPress(ByVal KeyAscii As Integer)
    RaiseEvent OnKeyPress(KeyAscii)
    If KeyAscii = Asc("S") Or KeyAscii = 19 Then SaveScreenShot
End Sub

Private Sub Disp_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim Tmp As Integer
    
    On Error GoTo DownErr
    
    If Not bHandleEvents Then Exit Sub
    SetCursor 0
    ButtonPress = OverButton(MouseX, MouseY)
    hDC = Disp.GetDC
    RestoreMouseBack
    If ButtonPress >= 0 Then
        DrawButton ButtonPress, True
    Else
        Tmp = OverPalCell(MouseX, MouseY)
        If Tmp >= 0 Then
            If Button And 1 Then
                CurPalEntry = Tmp
            End If
            If Button And 2 Then
                SecPalEntry = Tmp
            End If
            DrawPalette
        End If
        GridPress MouseX, MouseY, 1, Button
    End If
    Disp.ReleaseDC
    hDC = 0
    TileSetPress MouseX, MouseY
    hDC = Disp.GetDC
    SaveMouseBack
    Disp.ReleaseDC
    hDC = 0
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
    DrawMouse
    Disp.Flip
    Exit Sub
    
DownErr:
    Disp.Close
    MsgBox Err.Description
End Sub

Private Sub Disp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim XPix As Integer, YPix As Integer
    
    On Error GoTo MouseMoveErr
    
    If Not bHandleEvents Then Exit Sub
    SetCursor 0
    If hDC Then
        Disp.Close
        Err.Raise vbObjectError, , "Device context already locked"
    End If
    hDC = Disp.GetDC
    RestoreMouseBack
    MouseX = X
    MouseY = Y
    If Button Then
        GridPress MouseX, MouseY, 2, Button
    Else
        If CurTool = 12 Or CurTool = 13 Then
            If CalcPix(MouseX, MouseY, XPix, YPix) Then
                If bTempImage Then
                    RestoreGrid
                    Disp.PaintPicture TilePic, ImageLeft, ImageTop
                Else
                    StoreGrid
                    bTempImage = True
                End If
                If CurTool = 12 Then
                    SetClipping 1
                    StretchBlt hDC, XPix * Magnify + GridLeft, YPix * Magnify + GridTop, Disp.HimetricToPixelX(ClipPic.Width) * Magnify, Disp.HimetricToPixelY(ClipPic.Height) * Magnify, hDC, ClipLeft, ClipTop, Disp.HimetricToPixelX(ClipPic.Width), Disp.HimetricToPixelY(ClipPic.Height), SRCCOPY
                    SetClipping 0
                    PressThrough False, True
                End If
                If CurTool = 13 Then
                    PressThrough False, True
                    RotateClip GridLeft + XPix * Magnify, GridTop + YPix * Magnify, (Button And 2) > 0, 1, True
                End If
            Else
                If bTempImage Then
                    Disp.PaintPicture TilePic, ImageLeft, ImageTop
                    RestoreGrid
                    PressThrough
                End If
                bTempImage = False
            End If
        End If
    End If
    DrawCoord
    UpdateCurButton
    SaveMouseBack
    Disp.ReleaseDC
    hDC = 0
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
    DrawMouse
    Disp.Flip
    Exit Sub

MouseMoveErr:
    Disp.Close
    MsgBox Err.Description
End Sub

Sub UpdateCurButton()
Attribute UpdateCurButton.VB_Description = "If the mouse was pressed on a button, draw a pressed button if the mouse is still over it, otherwise draw an up button."
    If ButtonPress >= LBound(Buttons) Then
        If MouseX > Buttons(ButtonPress).rcFrame.Left And _
           MouseX < Buttons(ButtonPress).rcFrame.Right And _
           MouseY > Buttons(ButtonPress).rcFrame.Top And _
           MouseY < Buttons(ButtonPress).rcFrame.Bottom Then
            DrawButton ButtonPress, True
        Else
            DrawButton ButtonPress, False
        End If
    End If
End Sub

Sub SaveMouseBack()
Attribute SaveMouseBack.VB_Description = "Store the graphics behind the mouse pointer."
    If hDC = 0 Then
        Disp.Close
        Err.Raise vbObjectError, , "No device context"
    End If
    If MouseX <= GridLeft + Tileset.TileWidth * Magnify And _
       MouseX >= GridLeft And _
       MouseY <= GridTop + Tileset.TileHeight * Magnify And _
       MouseY >= GridTop Then
        BitBlt hdcMouseBack, 0, 0, 17, 17, hDC, MouseX - 8, MouseY - 8, SRCCOPY
    Else
        BitBlt hdcMouseBack, 0, 0, 17, 17, hDC, MouseX - 1, MouseY - 1, SRCCOPY
    End If
End Sub

Sub RestoreMouseBack()
Attribute RestoreMouseBack.VB_Description = "Restore the graphics behind the mouse pointer."
    Dim hdcFore As Long

    If hDC = 0 Then
        Disp.Close
        Err.Raise vbObjectError, , "No device context"
    End If
    
    hdcFore = Disp.GetDC(1)
    
    If MouseX <= GridLeft + Tileset.TileWidth * Magnify And _
       MouseX >= GridLeft And _
       MouseY <= GridTop + Tileset.TileHeight * Magnify And _
       MouseY >= GridTop Then
        BitBlt hDC, MouseX - 8, MouseY - 8, 17, 17, hdcMouseBack, 0, 0, SRCCOPY
        BitBlt hdcFore, MouseX - 8, MouseY - 8, 17, 17, hdcMouseBack, 0, 0, SRCCOPY
    Else
        BitBlt hDC, MouseX - 1, MouseY - 1, 17, 17, hdcMouseBack, 0, 0, SRCCOPY
        BitBlt hdcFore, MouseX - 1, MouseY - 1, 17, 17, hdcMouseBack, 0, 0, SRCCOPY
    End If
    
    Disp.ReleaseDC 1
    
End Sub

Sub DrawMouse()
Attribute DrawMouse.VB_Description = "Draw the mouse cursor."
    If hDC <> 0 Then
        Disp.Close
        Err.Raise vbObjectError, , "Device context locked"
    End If
    If MouseX <= GridLeft + Tileset.TileWidth * Magnify And _
       MouseX >= GridLeft And _
       MouseY <= GridTop + Tileset.TileHeight * Magnify And _
       MouseY >= GridTop Then
        Disp.DrawTile MouseTS, 1, MouseX - 8, MouseY - 8, 0, 0, 640, 480, True
    Else
        Disp.DrawTile MouseTS, 2, MouseX - 1, MouseY - 1, 0, 0, 640, 480, True
    End If
End Sub

Sub DrawAll()
Attribute DrawAll.VB_Description = "Draw the entire display."
    Dim I As Integer

    If hDC Then
        Disp.Close
        Err.Raise vbObjectError, , "Device context already locked"
    End If
    
    SetCursor 0
    Disp.Cls &HC0C0C0
    
    If Not (TilePic Is Nothing) Then
        Disp.PaintPicture TilePic, ImageLeft, ImageTop
    End If
    
    hDC = Disp.GetDC

    DrawGrid
    DrawPalette
    DrawCoord
    DrawAngle
    
    For I = LBound(Buttons) To UBound(Buttons)
        If I = ButtonPress Then
            If MouseX > Buttons(ButtonPress).rcFrame.Left And _
               MouseX < Buttons(ButtonPress).rcFrame.Right And _
               MouseY > Buttons(ButtonPress).rcFrame.Top And _
               MouseY < Buttons(ButtonPress).rcFrame.Bottom Then
                DrawButton ButtonPress, True
            Else
                DrawButton ButtonPress, False
            End If
        Else
            DrawButton I, False
        End If
    Next I
    Disp.ReleaseDC
    hDC = 0
    
    DrawClip
    DrawTileset
    
    hDC = Disp.GetDC
    SaveMouseBack
    Disp.ReleaseDC
    hDC = 0
    
    DrawMouse
    CopyBuf False

End Sub

Sub DrawButton(ByVal Index As Integer, ByVal State As Boolean)
Attribute DrawButton.VB_Description = "Draw the specified button in the specified pushed state."
    Dim bTempDC As Boolean
    Dim hOldFont As Long
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    DrawFrameControl hDC, Buttons(Index).rcFrame, DFC_BUTTON, DFCS_BUTTONPUSH Or (DFCS_PUSHED And State)
    hOldFont = SelectObject(hDC, ButtonFont.hFont)
    SetBkMode hDC, Transparent
    DrawText hDC, Buttons(Index).Caption, Len(Buttons(Index).Caption), Buttons(Index).rcFrame, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    SelectObject hDC, hOldFont
    
    With Buttons(Index).rcFrame
        CopyBuf False, .Left, .Top, .Right - .Left + 1, .Bottom - .Top + 1
    End With
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub CalcButtons()
Attribute CalcButtons.VB_Description = "Calculate appropriate sizes for all the buttons based on their text."
    Dim bTempDC As Boolean
    Dim I As Integer
    Dim cbCaption As Size
    Dim hOldFont As Long
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    hOldFont = SelectObject(hDC, ButtonFont.hFont)
    
    For I = LBound(Buttons) To UBound(Buttons)
        GetTextExtentPoint32 hDC, Buttons(I).Caption, Len(Buttons(I).Caption), cbCaption
        Buttons(I).rcFrame.Right = Buttons(I).rcFrame.Left + cbCaption.cx + 8
        Buttons(I).rcFrame.Bottom = Buttons(I).rcFrame.Top + cbCaption.cy + 4
    Next I
    
    SelectObject hDC, hOldFont
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
End Sub

Private Sub Disp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo UpErr
    
    If Not bHandleEvents Then Exit Sub
    SetCursor 0
    If OverButton(MouseX, MouseY) = ButtonPress And ButtonPress >= LBound(Buttons) Then
        TriggerButton ButtonPress
    End If
    hDC = Disp.GetDC
    RestoreMouseBack
    If ButtonPress >= 0 Then
        DrawButton ButtonPress, False
        ButtonPress = -1
    Else
        GridPress MouseX, MouseY, 3, Button
    End If
    SaveMouseBack
    Disp.ReleaseDC
    hDC = 0
    If CurTool = 11 Then
        If Not (ClipPic Is Nothing) Then
            DrawClip
            CopyBuf False, ClipLeft, ClipTop, Disp.HimetricToPixelX(ClipPic.Width), Disp.HimetricToPixelY(ClipPic.Height)
        End If
    End If
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
    DrawMouse
    Disp.Flip
    Exit Sub
    
UpErr:
    Disp.Close
    MsgBox Err.Description
End Sub

Sub TriggerButton(ByVal Index As Integer)
Attribute TriggerButton.VB_Description = "Handle the action for a button."
    Dim Red As Integer, Green As Integer, Blue As Integer
    Dim Red1 As Byte, Green1 As Byte, Blue1 As Byte
    Dim Red2 As Byte, Green2 As Byte, Blue2 As Byte
    Dim I As Integer

    If Index >= 1 And Index <= 12 Then 'Color change
        ColorToRGB PalEntries(CurPalEntry), Red, Green, Blue
    End If
    
    Select Case Index
    Case 0 ' Close
        RaiseEvent OnEditComplete
        If Not GameHost Is Nothing Then
            GameHost.RunStartScript
            If GameHost.CheckForError Then Exit Sub
        End If
    Case 1 ' R--
        Red = Red - 20
    Case 2 ' R-
        Red = Red - 1
    Case 3 ' R+
        Red = Red + 1
    Case 4 ' R++
        Red = Red + 20
    Case 5 ' G--
        Green = Green - 20
    Case 6 ' G-
        Green = Green - 1
    Case 7 ' G+
        Green = Green + 1
    Case 8 ' G++
        Green = Green + 20
    Case 9 ' B--
        Blue = Blue - 20
    Case 10 ' B-
        Blue = Blue - 1
    Case 11 ' B+
        Blue = Blue + 1
    Case 12 ' B++
        Blue = Blue + 20
    Case 13 ' Gradient
        If CurPalEntry = SecPalEntry Then Exit Sub
        ColorToRGBByte PalEntries(CurPalEntry), Red1, Green1, Blue1
        ColorToRGBByte PalEntries(SecPalEntry), Red2, Green2, Blue2
        For I = CurPalEntry To SecPalEntry Step Sgn(SecPalEntry - CurPalEntry)
            Red = (Red1 * Abs(I - SecPalEntry) + Red2 * Abs(I - CurPalEntry)) / Abs(SecPalEntry - CurPalEntry)
            Green = (Green1 * Abs(I - SecPalEntry) + Green2 * Abs(I - CurPalEntry)) / Abs(SecPalEntry - CurPalEntry)
            Blue = (Blue1 * Abs(I - SecPalEntry) + Blue2 * Abs(I - CurPalEntry)) / Abs(SecPalEntry - CurPalEntry)
            PalEntries(I) = RGB(Red, Green, Blue)
        Next
        DrawAll
    Case 14 ' Reset
        InitPalette
        DrawAll
    Case 15 ' Pixel
        CurTool = 0
    Case 16 ' Pencil
        CurTool = 1
    Case 17 ' Line
        CurTool = 2
    Case 18 ' Vertical Gradient
        CurTool = 3
    Case 19 ' Horizontal Gradient
        CurTool = 4
    Case 20 ' Empty Rectangle
        CurTool = 5
    Case 21 ' Solid Rectangle
        CurTool = 6
    Case 22 ' Flood Fill
        CurTool = 7
    Case 23 ' Hollow Ellipse
        CurTool = 8
    Case 24 ' Solid Ellipse
        CurTool = 9
    Case 25 ' Color Pick
        CurTool = 10
        DrawPalette
    Case 26 ' Make R.Button=Pick / Draw
        If RButtonPick Then
            Buttons(Index).Caption = "Make R. Button Pick"
            RButtonPick = False
        Else
            Buttons(Index).Caption = "Make R. Button Draw"
            RButtonPick = True
        End If
        CalcButtons
        DrawAll
    Case 27 ' Copy
        CurTool = 11
    Case 28 ' Paste
        CurTool = 12
    Case 29 ' Paste + (rotate, transparent, translucent)
        CurTool = 13
    Case 30 ' Rotate Left 15°
        RotAngle = RotAngle + 15
        DrawAngle
    Case 31 ' Rotate Left 1°
        RotAngle = RotAngle + 1
        DrawAngle
    Case 32 ' Rotate Right 1°
        RotAngle = RotAngle - 1
        DrawAngle
    Case 33 ' Rotate Right 15°
        RotAngle = RotAngle - 15
        DrawAngle
    Case 34 ' Translucency
        Translucent = (Translucent + 10) Mod 100
        Buttons(Index).Caption = "Translucency " & CStr(Translucent) & "%"
        CalcButtons
        DrawAll
    Case 35 ' AA On/Off
        If AAOn Then
            Buttons(Index).Caption = "Turn AA On"
            AAOn = False
        Else
            Buttons(Index).Caption = "Turn AA Off"
            AAOn = True
        End If
        CalcButtons
        DrawAll
    Case 36 ' X Flip
        FlipClip 1
    Case 37 ' X Flip
        FlipClip 2
    Case 38 ' Undo
        Undo
    Case 39 ' Redo
        Redo
    Case 40 ' Clear
        MakeBlankTile Tileset.TileWidth, Tileset.TileHeight
        DrawAll
        PushUndo
    Case 41 ' Left Shift
        PicShift -1, 0
        DrawGrid
        PressThrough False, True
    Case 42 ' Right Shift
        PicShift 1, 0
        DrawGrid
        PressThrough False, True
    Case 43 ' Up Shift
        PicShift 0, -1
        DrawGrid
        PressThrough False, True
    Case 44 ' Down Shift
        PicShift 0, 1
        DrawGrid
        PressThrough False, True
    Case 45 ' First
        CurCell = 0
        DrawTileset
    Case 46 ' Previous
        If CurCell > 0 Then CurCell = CurCell - 1
        DrawTileset
    Case 47 ' Next
        If CurCell < Tileset.TileColumns * TileRows - 1 Then CurCell = CurCell + 1
        DrawTileset
    Case 48 ' Last
        CurCell = Tileset.TileColumns * TileRows - 1
        DrawTileset
    Case 49 ' Get
        LoadTile CurCell
        PushUndo
        DrawGrid
        PressThrough
    Case 50 ' Store
        SaveTile CurCell
        DrawTileset
    Case 51 ' Smooth
        CurTool = 14
    Case 52 ' Make 16-bit
        Palette16Bit
    Case 53 ' Rotate 360° / 10°
        Rotate360by10
    Case 54 ' Color Change
        CurTool = 15
    Case 55 ' Grid
        bGrid = Not bGrid
        DrawGrid
        PressThrough False, True
    End Select
        
    If Index >= 1 And Index <= 12 Then 'Color change
        If Red < 0 Then Red = 0
        If Red > 255 Then Red = 255
        If Green < 0 Then Green = 0
        If Green > 255 Then Green = 255
        If Blue < 0 Then Blue = 0
        If Blue > 255 Then Blue = 255
        PalEntries(CurPalEntry) = RGB(Red, Green, Blue)
        DrawPalette
    End If
    
End Sub

Private Sub Rotate360by10()
    Dim I As Integer
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC()
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    Set ClipPic = CapturePicture(hDC, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight)
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
    SaveTile CurCell
    DrawAll
    For I = 1 To 35
        RotAngle = I * 10
        MakeBlankTile Tileset.TileWidth, Tileset.TileHeight
        Disp.PaintPicture TilePic, ImageLeft, ImageTop
        PressThrough True
        RotateClip ImageLeft + Int(Tileset.TileWidth / 2), ImageTop + Int(Tileset.TileHeight / 2), False, 2
        SaveTile (I + CurCell) Mod (Tileset.TileColumns * Disp.HimetricToPixelY(Tiles.Height) / Tileset.TileHeight)
    Next I
End Sub

Private Function OverButton(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim I As Integer
    
    OverButton = -1
    For I = LBound(Buttons) To UBound(Buttons)
        If X > Buttons(I).rcFrame.Left And _
           X < Buttons(I).rcFrame.Right And _
           Y > Buttons(I).rcFrame.Top And _
           Y < Buttons(I).rcFrame.Bottom Then
            OverButton = I
        End If
    Next
    
End Function

Sub InitPalette()
Attribute InitPalette.VB_Description = "Set up the initial set of colors."
    Dim I As Integer

    PalCellWidth = 10
    PalCellHeight = 10
    PalRows = 2
    PalCols = 30
    PalTop = 480 - PalRows * PalCellHeight - 3
    PalLeft = 0
    ReDim PalEntries(PalRows * PalCols - 1)
    For I = LBound(PalEntries) To UBound(PalEntries)
        PalEntries(I) = &HFFFFFF / (UBound(PalEntries) - LBound(PalEntries)) * (I - LBound(PalEntries))
    Next I
End Sub

Sub InitButtons()
Attribute InitButtons.VB_Description = "Create all the buttons in their initial locations with the initial text."
    Dim B As Integer

    With Buttons(0)
        .rcFrame.Left = 500
        .rcFrame.Top = 440
        .Caption = "Close"
    End With
    For B = 1 To 12
        With Buttons(B).rcFrame
            .Left = ((B - 1) \ 4) * 32
            .Top = PalTop - 125 + ((B - 1) Mod 4) * 25
        End With
    Next
    
    Buttons(1).Caption = "R--"
    Buttons(2).Caption = "R-"
    Buttons(3).Caption = "R+"
    Buttons(4).Caption = "R++"
    Buttons(5).Caption = "G--"
    Buttons(6).Caption = "G-"
    Buttons(7).Caption = "G+"
    Buttons(8).Caption = "G++"
    Buttons(9).Caption = "B--"
    Buttons(10).Caption = "B-"
    Buttons(11).Caption = "B+"
    Buttons(12).Caption = "B++"
    
    With Buttons(13)
        .rcFrame.Left = 0
        .rcFrame.Top = PalTop - 25
        .Caption = "Gradient"
    End With
    With Buttons(14)
        .rcFrame.Left = 70
        .rcFrame.Top = PalTop - 25
        .Caption = "Reset"
    End With
    
    For B = 15 To 29
        With Buttons(B).rcFrame
            .Left = 500
            .Top = (B - 15) * 25
        End With
    Next
    
    Buttons(15).Caption = "Pixel"
    Buttons(16).Caption = "Pencil"
    Buttons(17).Caption = "Line"
    Buttons(18).Caption = "Vertical Gradient"
    Buttons(19).Caption = "Horizontal Gradient"
    Buttons(20).Caption = "Empty Rectangle"
    Buttons(21).Caption = "Solid Rectangle"
    Buttons(22).Caption = "Flood Fill"
    Buttons(23).Caption = "Hollow Ellipse"
    Buttons(24).Caption = "Solid Ellipse"
    Buttons(25).Caption = "Color Pick"
    Buttons(26).Caption = "Make R. Button Pick"
    Buttons(27).Caption = "Copy"
    Buttons(28).Caption = "Paste"
    Buttons(29).Caption = "Paste +"
    
    For B = 30 To 35
        With Buttons(B).rcFrame
            .Top = 100 + 25 * (B - 30)
            .Left = 330
        End With
    Next
    
    Buttons(30).Caption = "Left 15°"
    Buttons(31).Caption = "Left 1°"
    Buttons(32).Caption = "Right 1°"
    Buttons(33).Caption = "Right 15°"
    Buttons(34).Caption = "Translucency " & CStr(Translucent) & "%"
    Buttons(35).Caption = "Turn AA On"
    
    For B = 36 To 37
        With Buttons(B).rcFrame
            .Top = 175 + 25 * (B - 36)
            .Left = 270
        End With
    Next
    
    Buttons(36).Caption = "X Flip"
    Buttons(37).Caption = "Y Flip"
    
    For B = 38 To 44
        With Buttons(B).rcFrame
            .Top = (B - 38) * 25
            .Left = 400
        End With
    Next
    
    Buttons(38).Caption = "Undo"
    Buttons(39).Caption = "Redo"
    Buttons(40).Caption = "Clear"
    Buttons(41).Caption = "L. Shift"
    Buttons(42).Caption = "R. Shift"
    Buttons(43).Caption = "U. Shift"
    Buttons(44).Caption = "D. Shift"
    
    With Buttons(45)
        .rcFrame.Left = 1
        .rcFrame.Top = 280
        .Caption = "|<<"
    End With
    With Buttons(46)
        .rcFrame.Left = 30
        .rcFrame.Top = 280
        .Caption = "<"
    End With
    With Buttons(47)
        .rcFrame.Left = 50
        .rcFrame.Top = 280
        .Caption = ">"
    End With
    With Buttons(48)
        .rcFrame.Left = 70
        .rcFrame.Top = 280
        .Caption = ">>|"
    End With
    With Buttons(49)
        .rcFrame.Left = 105
        .rcFrame.Top = 280
        .Caption = "Get"
    End With
    With Buttons(50)
        .rcFrame.Left = 140
        .rcFrame.Top = 280
        .Caption = "Store"
    End With
    With Buttons(51)
        .rcFrame.Left = 500
        .rcFrame.Top = 375
        .Caption = "Smooth"
    End With
    With Buttons(52)
        .rcFrame.Left = 140
        .rcFrame.Top = PalTop - 25
        .Caption = "Make 16-bit"
    End With
    With Buttons(53)
        .rcFrame.Left = 330
        .rcFrame.Top = 250
        .Caption = "Make 360°/10°"
    End With
    With Buttons(54)
        .rcFrame.Left = 500
        .rcFrame.Top = 400
        .Caption = "Color Change"
    End With
    With Buttons(55)
        .rcFrame.Left = 190
        .rcFrame.Top = 280
        .Caption = "Grid"
    End With
    ButtonPress = -1
End Sub

Sub DrawPalette()
Attribute DrawPalette.VB_Description = "Draw the color palette."
    Dim Y As Long
    Dim X As Long
    Dim hBrush As Long
    Dim rcEntry As RECT
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    With rcEntry
        .Left = PalLeft
        .Top = PalTop
        .Right = .Left + PalCellWidth * PalCols + 1
        .Bottom = .Top + PalCellHeight * PalRows + 1
    End With
    FillRect hDC, rcEntry, GetStockObject(BLACK_BRUSH)
    
    For Y = 0 To PalRows - 1
        For X = 0 To PalCols - 1
            If X + Y * PalCols = CurPalEntry Or X + Y * PalCols = SecPalEntry Then
                With rcEntry
                    .Left = X * PalCellWidth + PalLeft
                    .Top = Y * PalCellHeight + PalTop
                    .Right = .Left + PalCellWidth + 1
                    .Bottom = .Top + PalCellWidth + 1
                End With
                If X + Y * PalCols = CurPalEntry Then
                    FillRect hDC, rcEntry, GetStockObject(WHITE_BRUSH)
                Else
                    FillRect hDC, rcEntry, GetStockObject(GRAY_BRUSH)
                End If
            End If
            With rcEntry
                .Left = X * PalCellWidth + PalLeft + 1
                .Top = Y * PalCellHeight + PalTop + 1
                .Right = .Left + PalCellWidth - 1
                .Bottom = .Top + PalCellWidth - 1
            End With
            hBrush = CreateSolidBrush(PalEntries(X + Y * PalCols))
            FillRect hDC, rcEntry, hBrush
            DeleteObject hBrush
        Next
    Next
    
    DrawRGB
    
    CopyBuf False, PalLeft, PalTop, PalCols * PalCellWidth + 1, PalRows * PalCellHeight + 1
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Function OverPalCell(ByVal X As Integer, ByVal Y As Integer) As Integer
Attribute OverPalCell.VB_Description = "Returns the palette index under the mouse, -1 if none."
    Dim I As Integer
    
    If X < PalLeft + PalCols * PalCellWidth And _
       Y < PalTop + PalRows * PalCellHeight And _
       X > PalLeft And Y > PalTop Then
        OverPalCell = (X - PalLeft) \ PalCellWidth + ((Y - PalTop) \ PalCellHeight) * PalCols
    Else
        OverPalCell = -1
    End If
End Function

Sub GridPress(ByVal X As Integer, ByVal Y As Integer, ByVal Action As Integer, ByVal Button As Integer)
Attribute GridPress.VB_Description = "Handle a mouse action on the image. Action may be 1=MouseDown, 2=MouseMove or 3=MouseUp."
    Dim XPix As Integer, YPix As Integer
    Dim bTempDC As Boolean
    Static XStart As Integer, YStart As Integer, bActStarted As Boolean
    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    Dim UseEntry As Integer, OtherEntry As Integer
    Dim Red As Byte, Green As Byte, Blue As Byte
    Dim Red1 As Byte, Green1 As Byte, Blue1 As Byte
    Dim Red2 As Byte, Green2 As Byte, Blue2 As Byte
    Dim I As Integer, J As Integer
    Dim IMin As Integer, IMax As Integer
    Dim JMin As Integer, JMax As Integer
    Dim rcRect As RECT
    
    On Error GoTo GridPressErr
    
    If Button And 1 Then
        UseEntry = CurPalEntry
        OtherEntry = SecPalEntry
    Else
        UseEntry = SecPalEntry
        OtherEntry = CurPalEntry
    End If
    
    ColorToRGBByte PalEntries(UseEntry), Red1, Green1, Blue1
    ColorToRGBByte PalEntries(OtherEntry), Red2, Green2, Blue2
    
    If Not CalcPix(X, Y, XPix, YPix) Then Exit Sub

    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    If RButtonPick And ((Button And 2) > 0) And CurTool <> 10 Then
        PalEntries(CurPalEntry) = GetPixel(hDC, XPix + ImageLeft, YPix + ImageTop)
        DrawPalette
        GoTo SkipTool
    End If
    
    If Action = 1 Then bActStarted = True
    If Action = 3 Then bActStarted = False
    If Action = 2 And Not (bActStarted) Then Action = 1: bActStarted = True
    
    Select Case CurTool
    Case 0
        SetPixel hDC, XPix + ImageLeft, YPix + ImageTop, PalEntries(UseEntry)
        DrawMagPixel XPix, YPix
        PressThrough True, False
        If Action = 3 Then PushUndo
    Case 1 ' Pencil
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            SetPixel hDC, XPix + ImageLeft, YPix + ImageTop, PalEntries(UseEntry)
            DrawMagPixel XPix, YPix
        Case 2
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            MoveToEx hDC, XStart + ImageLeft, YStart + ImageTop, 0
            LineTo hDC, XPix + ImageLeft, YPix + ImageTop
            SelectObject hDC, hOldPen
            DeleteObject hPen
            hPen = CreatePen(PS_SOLID, Magnify, PalEntries(UseEntry))
            SelectObject hDC, hPen
            MoveToEx hDC, XStart * Magnify + GridLeft + Magnify \ 2, YStart * Magnify + GridTop + Magnify \ 2, 0
            LineTo hDC, XPix * Magnify + GridLeft + Magnify \ 2, YPix * Magnify + GridTop + Magnify \ 2
            SelectObject hDC, hOldPen
            DeleteObject hPen
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            With rcRect
                CopyBuf False, .Left, .Top, .Right - .Left + 1, .Bottom - .Top + 1
            End With
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            With rcRect
                CopyBuf False, .Left, .Top, .Right - .Left + 1, .Bottom - .Top + 1
            End With
            XStart = XPix: YStart = YPix
        Case 3
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            MoveToEx hDC, XStart + ImageLeft, YStart + ImageTop, 0
            LineTo hDC, XPix + ImageLeft, YPix + ImageTop
            SelectObject hDC, hOldPen
            DeleteObject hPen
            DrawGrid
            PushUndo
        End Select
    Case 2 ' Line
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hPen = CreatePen(PS_SOLID, Magnify, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            MoveToEx hDC, XStart * Magnify + GridLeft + Magnify / 2, YStart * Magnify + GridTop + Magnify / 2, 0
            LineTo hDC, XPix * Magnify + GridLeft + Magnify / 2, YPix * Magnify + GridTop + Magnify / 2
            SelectObject hDC, hOldPen
            DeleteObject hPen
        Case 3
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            MoveToEx hDC, XStart + ImageLeft, YStart + ImageTop, 0
            LineTo hDC, XPix + ImageLeft + Sgn(XPix - XStart), YPix + ImageTop + Sgn(YPix - YStart)
            SelectObject hDC, hOldPen
            DeleteObject hPen
            DrawGrid
            PushUndo
        End Select
    Case 3 ' Vertical Gradient
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hOldPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
            If YStart <> YPix Then
                For I = YStart To YPix Step Sgn(YPix - YStart)
                    Red = (Red1 * Abs(I - YPix) + Red2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    Green = (Green1 * Abs(I - YPix) + Green2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    Blue = (Blue1 * Abs(I - YPix) + Blue2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    hPen = CreatePen(PS_SOLID, Magnify, RGB(Red, Green, Blue))
                    SelectObject hDC, hPen
                    MoveToEx hDC, XStart * Magnify + GridLeft + Magnify / 2, I * Magnify + GridTop + Magnify / 2, 0
                    LineTo hDC, XPix * Magnify + GridLeft + Magnify / 2, I * Magnify + GridTop + Magnify / 2
                    SelectObject hDC, hOldPen
                    DeleteObject hPen
                Next
            End If
        Case 3
            hOldPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
            If YStart <> YPix Then
                For I = YStart To YPix Step Sgn(YPix - YStart)
                    Red = (Red1 * Abs(I - YPix) + Red2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    Green = (Green1 * Abs(I - YPix) + Green2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    Blue = (Blue1 * Abs(I - YPix) + Blue2 * Abs(I - YStart)) / Abs(YPix - YStart)
                    hPen = CreatePen(PS_SOLID, 1, RGB(Red, Green, Blue))
                    SelectObject hDC, hPen
                    MoveToEx hDC, XStart + ImageLeft, I + ImageTop, 0
                    LineTo hDC, XPix + ImageLeft + Sgn(XPix - XStart), I + ImageTop
                    SelectObject hDC, hOldPen
                    DeleteObject hPen
                Next
            End If
            DrawGrid
            PushUndo
        End Select
    Case 4 ' Horizontal Gradient
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hOldPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
            If XStart <> XPix Then
                For I = XStart To XPix Step Sgn(XPix - XStart)
                    Red = (Red1 * Abs(I - XPix) + Red2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    Green = (Green1 * Abs(I - XPix) + Green2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    Blue = (Blue1 * Abs(I - XPix) + Blue2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    hPen = CreatePen(PS_SOLID, Magnify, RGB(Red, Green, Blue))
                    SelectObject hDC, hPen
                    MoveToEx hDC, I * Magnify + GridLeft + Magnify / 2, YStart * Magnify + GridTop + Magnify / 2, 0
                    LineTo hDC, I * Magnify + GridLeft + Magnify / 2, YPix * Magnify + GridTop + Magnify / 2
                    SelectObject hDC, hOldPen
                    DeleteObject hPen
                Next
            End If
        Case 3
            hOldPen = SelectObject(hDC, GetStockObject(WHITE_PEN))
            If XStart <> XPix Then
                For I = XStart To XPix Step Sgn(XPix - XStart)
                    Red = (Red1 * Abs(I - XPix) + Red2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    Green = (Green1 * Abs(I - XPix) + Green2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    Blue = (Blue1 * Abs(I - XPix) + Blue2 * Abs(I - XStart)) / Abs(XPix - XStart)
                    hPen = CreatePen(PS_SOLID, 1, RGB(Red, Green, Blue))
                    SelectObject hDC, hPen
                    MoveToEx hDC, I + ImageLeft, YStart + ImageTop, 0
                    LineTo hDC, I + ImageLeft, YPix + ImageTop + Sgn(YPix - YStart)
                    SelectObject hDC, hOldPen
                    DeleteObject hPen
                Next
            End If
            DrawGrid
            PushUndo
        End Select
    Case 5 ' Empty Rectangle
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
            hPen = CreatePen(PS_SOLID, Magnify, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True, False)
            Rectangle hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
        Case 3
            hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            Rectangle hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DrawGrid
            PushUndo
        End Select
    Case 6 ' Solid Rectangle
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            hBrush = CreateSolidBrush(PalEntries(UseEntry))
            FillRect hDC, rcRect, hBrush
            DeleteObject hBrush
        Case 3
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            hBrush = CreateSolidBrush(PalEntries(UseEntry))
            FillRect hDC, rcRect, hBrush
            DeleteObject hBrush
            DrawGrid
            PushUndo
        End Select
    Case 7 ' Flood Fill
        If Action = 3 Then
            hBrush = CreateSolidBrush(PalEntries(UseEntry))
            hOldBrush = SelectObject(hDC, hBrush)
            ExtFloodFill hDC, XPix + ImageLeft, YPix + ImageTop, GetPixel(hDC, XPix + ImageLeft, YPix + ImageTop), FLOODFILLSURFACE
            SelectObject hDC, hOldBrush
            DeleteObject hBrush
            DrawGrid
            PushUndo
        End If
    Case 8 ' Empty Ellipse
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
            hPen = CreatePen(PS_SOLID, Magnify, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            Ellipse hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
        Case 3
            hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            Ellipse hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DrawGrid
            PushUndo
        End Select
    Case 9 ' Solid Ellipse
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            hBrush = CreateSolidBrush(PalEntries(UseEntry))
            hOldBrush = SelectObject(hDC, hBrush)
            hPen = CreatePen(PS_SOLID, Magnify, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            Ellipse hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DeleteObject hBrush
        Case 3
            hBrush = CreateSolidBrush(PalEntries(UseEntry))
            hOldBrush = SelectObject(hDC, hBrush)
            hPen = CreatePen(PS_SOLID, 1, PalEntries(UseEntry))
            hOldPen = SelectObject(hDC, hPen)
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            Ellipse hDC, rcRect.Left, rcRect.Top, rcRect.Right, rcRect.Bottom
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DeleteObject hBrush
            DrawGrid
            PushUndo
        End Select
    Case 10 ' Color Pick
        PalEntries(UseEntry) = GetPixel(hDC, XPix + ImageLeft, YPix + ImageTop)
        DrawPalette
    Case 11 ' Copy
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            DrawFocusRect hDC, rcRect
        Case 3
            rcRect = MakeRect(XStart, YStart, XPix, YPix, False)
            With rcRect
                Set ClipPic = CapturePicture(hDC, .Left, .Top, .Right - .Left, .Bottom - .Top)
            End With
            DrawGrid
        End Select
    Case 12 ' Paste
        If Action = 1 Then BitBlt hDC, XPix + ImageLeft, YPix + ImageTop, Disp.HimetricToPixelX(ClipPic.Width), Disp.HimetricToPixelY(ClipPic.Height), hDC, ClipLeft, ClipTop, SRCCOPY
        DrawGrid
        bTempImage = False
        If Action = 3 Then PushUndo
    Case 13 ' Paste + (rotate, transparent, translucent)
        If Action = 1 Then RotateClip ImageLeft + XPix, ImageTop + YPix, (Button And 2) > 0, 2
        DrawGrid
        bTempImage = False
        If Action = 3 Then PushUndo
    Case 14 ' Smooth
        Select Case Action
        Case 1
            XStart = XPix: YStart = YPix
            StoreGrid
        Case 2
            RestoreGrid
            rcRect = MakeRect(XStart, YStart, XPix, YPix, True)
            DrawFocusRect hDC, rcRect
        Case 3
            If YPix <> YStart And XPix <> XStart Then
                For J = YStart To YPix Step Sgn(YPix - YStart)
                    For I = XStart To XPix Step Sgn(XPix - XStart)
                        IMin = IIf(I > XStart, I - 1, I)
                        IMax = IIf(I < XPix, I + 1, I)
                        JMin = IIf(J > YStart, J - 1, J)
                        JMax = IIf(J < YPix, J + 1, J)
                        SetPixelV hDC, ImageLeft + I, ImageTop + J, _
                            Mix9Colors(GetPixel(hDC, IMin * Magnify + GridLeft + 1, JMin * Magnify + GridTop + 1), GetPixel(hDC, I * Magnify + GridLeft + 1, JMin * Magnify + GridTop + 1), GetPixel(hDC, IMax * Magnify + GridLeft + 1, JMin * Magnify + GridTop + 1), _
                                GetPixel(hDC, IMin * Magnify + GridLeft + 1, J * Magnify + GridTop + 1), GetPixel(hDC, I * Magnify + GridLeft + 1, J * Magnify + GridTop + 1), GetPixel(hDC, IMax * Magnify + GridLeft + 1, J * Magnify + GridTop + 1), _
                                GetPixel(hDC, IMin * Magnify + GridLeft + 1, JMax * Magnify + GridTop + 1), GetPixel(hDC, I * Magnify + GridLeft + 1, JMax * Magnify + GridTop + 1), GetPixel(hDC, IMax * Magnify + GridLeft + 1, JMax * Magnify + GridTop + 1))
                    Next
                Next
                DrawGrid
                PushUndo
            End If
        End Select
    Case 15 ' Color Change
        If Action = 1 Then
            hBrush = GetPixel(hDC, ImageLeft + XPix, ImageTop + YPix)
            For J = ImageTop To ImageTop + Tileset.TileHeight - 1
                For I = ImageLeft To ImageLeft + Tileset.TileWidth - 1
                    If GetPixel(hDC, I, J) = hBrush Then
                        SetPixelV hDC, I, J, PalEntries(UseEntry)
                    End If
                Next
            Next
            DrawGrid
            PushUndo
        End If
    End Select

    If Action = 1 Or Action = 3 Then ' MouseDown or MouseUp
        RestoreMouseBack
        Set TilePic = CapturePicture(hDC, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight)
        Disp.PaintPicture TilePic, ImageLeft, ImageTop
        DrawGrid
        PressThrough True, True
    End If
    
SkipTool:

    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
    Exit Sub
    
GridPressErr:
    If hDC Then Disp.ReleaseDC
    hDC = 0
    Disp.Close
    MsgBox "Draw operation failed: " & Err.Description
    
End Sub

Private Sub Disp_TileSetLost(ByVal Index As Integer)
    On Error Resume Next
    If Disp.TileSets(Index) Is Tileset Then
        Tileset.Restore
        Tileset.PaintPicture Tiles, 0, 0
    End If
    If Disp.TileSets(Index) Is MouseTS Then
        MouseTS.Restore
        MouseTS.PaintPicture GetMouseBmp(), 0, 0
        DrawAll
    End If
End Sub

Private Function MakeRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal bMagnify As Boolean, Optional ByVal bOutside As Boolean = True) As RECT
    If X1 < X2 Then
        MakeRect.Left = X1 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridLeft, ImageLeft) + IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 + 1)
        MakeRect.Right = X2 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridLeft + Magnify, ImageLeft + 1) - IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 - 1)
    Else
        MakeRect.Left = X2 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridLeft, ImageLeft) + IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 + 1)
        MakeRect.Right = X1 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridLeft + Magnify, ImageLeft + 1) - IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 - 1)
    End If
    If Y1 < Y2 Then
        MakeRect.Top = Y1 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridTop, ImageTop) + IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 + 1)
        MakeRect.Bottom = Y2 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridTop + Magnify, ImageTop + 1) - IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 - 1)
    Else
        MakeRect.Top = Y2 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridTop, ImageTop) + IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 + 1)
        MakeRect.Bottom = Y1 * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, GridTop + Magnify, ImageTop + 1) - IIf(bOutside Or (Not bMagnify), 0, Magnify / 2 - 1)
    End If
End Function

Function CalcPix(ByVal X As Integer, ByVal Y As Integer, ByRef XPix As Integer, ByRef YPix As Integer) As Boolean
Attribute CalcPix.VB_Description = "Given an X,Y screen coordinate, translate it into X,Y pixel on the image, return false if not on the image."
    If X > GridLeft And X < GridLeft + Tileset.TileWidth * Magnify And _
       Y > GridTop And Y < GridTop + Tileset.TileHeight * Magnify Then
        XPix = (X - GridLeft) \ Magnify
        YPix = (Y - GridTop) \ Magnify
        CalcPix = True
    ElseIf X >= ImageLeft And X < ImageLeft + Tileset.TileWidth And _
           Y >= ImageTop And Y < ImageTop + Tileset.TileHeight Then
        XPix = X - ImageLeft
        YPix = Y - ImageTop
        CalcPix = True
    Else
        CalcPix = False
    End If
End Function

Sub DrawCoord()
Attribute DrawCoord.VB_Description = "Draw the values indicating the location of the mouse over the tile image."
    Dim XPix As Integer, YPix As Integer
    Dim rcRect As RECT
    Dim hBrushBack As Long
    
    If hDC = 0 Then
        Disp.Close
        MsgBox "No device context"
    End If
    
    With rcRect
        .Left = 0
        .Top = GridTop + Magnify * Tileset.TileHeight + 2
        .Right = 100
        .Bottom = .Top + 15
    End With
    SetBkColor hDC, &HC0C0C0
    
    If Not CalcPix(MouseX, MouseY, XPix, YPix) Then
        hBrushBack = CreateSolidBrush(&HC0C0C0)
        FillRect hDC, rcRect, hBrushBack
        DeleteObject hBrushBack
    Else
        DrawText hDC, CStr(XPix) & "," & CStr(YPix) & Space$(20) & Chr$(0), -1, rcRect, DT_SINGLELINE
    End If
End Sub

Sub DrawRGB()
Attribute DrawRGB.VB_Description = "Draw the RGB values of the currently selected color."
    Dim Red As Byte, Green As Byte, Blue As Byte
    Dim rcRect As RECT
    Dim hOldFont As Long
    
    If hDC = 0 Then
        Disp.Close
        MsgBox "No device context"
    End If
    
    hOldFont = SelectObject(hDC, ButtonFont.hFont)
    SetBkColor hDC, &HC0C0C0
    
    ColorToRGBByte PalEntries(CurPalEntry), Red, Green, Blue
    
    With rcRect
        .Left = 100
        .Top = PalTop - 125
        .Right = .Left + 150
        .Bottom = .Top + 30
    End With
    
    DrawText hDC, "Primary:" & vbCrLf & "R=" & CStr(Red) & ", G=" & CStr(Green) & ", B=" & CStr(Blue) & "  " & Chr$(0), -1, rcRect, 0
    CopyBuf False, rcRect.Left, rcRect.Top, 150, 30
    
    ColorToRGBByte PalEntries(SecPalEntry), Red, Green, Blue
    
    With rcRect
        .Top = PalTop - 95
        .Bottom = .Top + 30
    End With
    
    DrawText hDC, "Secondary:" & vbCrLf & "R=" & CStr(Red) & ", G=" & CStr(Green) & ", B=" & CStr(Blue) & "  " & Chr$(0), -1, rcRect, 0
    CopyBuf False, rcRect.Left, rcRect.Top, 150, 30
    
    SelectObject hDC, hOldFont
    
End Sub

Sub DrawClip()
Attribute DrawClip.VB_Description = "Draw the copied image (clipboard)."
    Dim hOldFont As Long
    Dim rcRect As RECT
    
    On Error GoTo DrawClipErr
    
    If hDC Then
        Disp.Close
        MsgBox "Device context locked"
        Exit Sub
    End If
    
    If ClipPic Is Nothing Then Exit Sub
    
    hDC = Disp.GetDC
    hOldFont = SelectObject(hDC, ButtonFont.hFont)
    SetBkColor hDC, &HC0C0C0
    With rcRect
        .Left = ClipLeft
        .Top = ClipTop - 15
        .Right = .Left + 50
        .Bottom = .Top + 15
    End With
    DrawText hDC, "Clip", 4, rcRect, DT_SINGLELINE
    CopyBuf False, rcRect.Left, rcRect.Top, 50, 15
    Disp.ReleaseDC
    hDC = 0

    Disp.PaintPicture ClipPic, ClipLeft, ClipTop
    Exit Sub

DrawClipErr:
    If hDC Then Disp.ReleaseDC
    hDC = 0
    Disp.Close
    MsgBox "DrawClip failed: " & Err.Description
End Sub

Sub RotateClip(ByVal XDest As Long, ByVal YDest As Long, ByVal bErase As Boolean, ByVal ClipArea As Integer, Optional ByVal bMagnify As Boolean = False)
Attribute RotateClip.VB_Description = "Rotate the copied image (clipboard) onto the display. ClipArea: 1=Grid, 2=un-magnified image."
    Dim Radian As Double
    Dim CenterX As Double
    Dim CenterY As Double
    Dim CurRad As Double
    Dim Dst As Double
    Dim RedTL As Integer, GreenTL As Integer, BlueTL As Integer, WeightTL As Double
    Dim RedTR As Integer, GreenTR As Integer, BlueTR As Integer, WeightTR As Double
    Dim RedBL As Integer, GreenBL As Integer, BlueBL As Integer, WeightBL As Double
    Dim RedBR As Integer, GreenBR As Integer, BlueBR As Integer, WeightBR As Double
    Dim PixColor As Long
    Dim A As Double, B As Double
    Dim C As Double, d As Double
    Dim FinalColor As Long
    Dim X As Double, Y As Double
    Dim TransColor As Long
    Dim RotWidth As Double, RotHeight As Double
    Dim rcPixel As RECT
    Dim hbrushPixel As Long
    Dim hdcFore As Long
    Dim hrgnClip As Long
    Dim hrgnPic As Long
    Dim DrawColor As Long
    Dim CPOX As Double ' Horizontal offset to center of pixel
    Dim CPOY As Double ' Vertical offset to center of pixel
    
    On Error GoTo RotateErr
    
    If ClipPic Is Nothing Then Exit Sub
    
    hdcFore = Disp.GetDC(1)
    TransColor = -1
        
    CalcRotSize Disp.HimetricToPixelX(ClipPic.Width), Disp.HimetricToPixelY(ClipPic.Height), RotWidth, RotHeight
    Select Case ClipArea
    Case 0
        hrgnClip = 0
        hrgnPic = 0
    Case 1 ' Grid
        hrgnClip = CreateRectRgn(GridLeft, GridTop, GridLeft + Tileset.TileWidth * Magnify, GridTop + Tileset.TileHeight * Magnify)
        hrgnPic = CreateRectRgn(ClipLeft, ClipTop, ClipLeft + Disp.HimetricToPixelX(ClipPic.Width), ClipTop + Disp.HimetricToPixelY(ClipPic.Height))
    Case 2 ' Image
        hrgnClip = CreateRectRgn(ImageLeft, ImageTop, ImageLeft + Tileset.TileWidth, ImageTop + Tileset.TileHeight)
        hrgnPic = CreateRectRgn(ClipLeft, ClipTop, ClipLeft + Disp.HimetricToPixelX(ClipPic.Width), ClipTop + Disp.HimetricToPixelY(ClipPic.Height))
    End Select

    CenterX = Disp.HimetricToPixelX(ClipPic.Width) / 2#
    CenterY = Disp.HimetricToPixelY(ClipPic.Height) / 2#
    Radian = CDbl(RotAngle) * Pi / 180#
    
    SelectClipRgn hdcFore, hrgnPic
    
    PixColor = GetPixel(hdcFore, ClipLeft, ClipTop)
    If PixColor = GetPixel(hdcFore, ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) - 1, ClipTop) Then
        If PixColor = GetPixel(hdcFore, ClipLeft, ClipTop + Disp.HimetricToPixelY(ClipPic.Height) - 1) Then
            If PixColor = GetPixel(hdcFore, ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) - 1, ClipTop + Disp.HimetricToPixelY(ClipPic.Height) - 1) Then
                TransColor = PixColor
            End If
        End If
    End If
    
    If Abs(Cos(Radian)) > Abs(Sin(Radian)) Then
        CPOX = 0.5 - (CenterX - Int(CenterX))
        CPOY = 0.5 - (CenterY - Int(CenterY))
    Else
        CPOX = 0.5 - (CenterY - Int(CenterY))
        CPOY = 0.5 - (CenterX - Int(CenterX))
    End If
    
    For Y = Int(-RotHeight / 2#) - 1 + CPOY To -Int(-RotHeight / 2#) + 1 + CPOY
        For X = Int(-RotWidth / 2#) - 1 + CPOX To -Int(-RotWidth / 2#) + 1 + CPOX
            If X <> 0 Then
                CurRad = Atn(-Y / X)
            Else
                CurRad = -(Pi / 2) * Sgn(Y)
            End If
            If X < 0 Then
                CurRad = Pi + CurRad
            ElseIf Y > 0 Then
                CurRad = Pi * 2 + CurRad
            End If
            Dst = Sqr(X * X + Y * Y)
            CurRad = CurRad - Radian
            A = ClipLeft + CenterX + Cos(CurRad) * Dst - 0.5
            B = ClipTop + CenterY - Sin(CurRad) * Dst - 0.5
            
            SelectClipRgn hdcFore, hrgnPic
            
            If AAOn Then
                If Int(A) < ClipLeft Or Int(A) >= ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) Or _
                   Int(B) < ClipTop Or Int(B) >= ClipTop + Disp.HimetricToPixelY(ClipPic.Height) Then
                    WeightTL = 0
                Else
                    PixColor = GetPixel(hdcFore, Int(A), Int(B))
                    ColorToRGB PixColor, RedTL, GreenTL, BlueTL
                    C = Abs(A - Int(A))
                    d = Abs(B - Int(B))
                    WeightTL = Sq2 - Sqr(C * C + d * d)
                End If
                
                If -Int(-A) < ClipLeft Or -Int(-A) >= ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) Or _
                   Int(B) < ClipTop Or Int(B) >= ClipTop + Disp.HimetricToPixelY(ClipPic.Height) Then
                   WeightTR = 0
                Else
                    PixColor = GetPixel(hdcFore, -Int(-A), Int(B))
                    ColorToRGB PixColor, RedTR, GreenTR, BlueTR
                    C = Abs(A - -Int(-A))
                    d = Abs(B - Int(B))
                    WeightTR = Sq2 - Sqr(C * C + d * d)
                End If
                
                If Int(A) < ClipLeft Or Int(A) >= ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) Or _
                   -Int(-B) < ClipTop Or -Int(-B) >= ClipTop + Disp.HimetricToPixelY(ClipPic.Height) Then
                    WeightBL = 0
                Else
                    PixColor = GetPixel(hdcFore, Int(A), -Int(-B))
                    ColorToRGB PixColor, RedBL, GreenBL, BlueBL
                    C = Abs(A - Int(A))
                    d = Abs(B - -Int(-B))
                    WeightBL = Sq2 - Sqr(C * C + d * d)
                End If
                
                If -Int(-A) < ClipLeft Or -Int(-A) >= ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) Or _
                   -Int(-B) < ClipTop Or -Int(-B) >= ClipTop + Disp.HimetricToPixelY(ClipPic.Height) Then
                    WeightBR = 0
                Else
                    PixColor = GetPixel(hdcFore, -Int(-A), -Int(-B))
                    ColorToRGB PixColor, RedBR, GreenBR, BlueBR
                    C = Abs(A - -Int(-A))
                    d = Abs(B - -Int(-B))
                    WeightBR = Sq2 - Sqr(C * C + d * d)
                End If
                            
                A = WeightTL + WeightTR + WeightBL + WeightBR
                If A > 0 Then
                    WeightTL = WeightTL / A
                    WeightTR = WeightTR / A
                    WeightBL = WeightBL / A
                    WeightBR = WeightBR / A
                End If
    
                FinalColor = RGB(RedTL * WeightTL + RedTR * WeightTR + RedBL * WeightBL + RedBR * WeightBR, _
                                 GreenTL * WeightTL + GreenTR * WeightTR + GreenBL * WeightBL + GreenBR * WeightBR, _
                                 BlueTL * WeightTL + BlueTR * WeightTR + BlueBL * WeightBL + BlueBR * WeightBR)
            Else
                C = CInt(A)
                d = CInt(B)
                FinalColor = GetPixel(hdcFore, C, d)
                If C < ClipLeft Or C >= ClipLeft + Disp.HimetricToPixelX(ClipPic.Width) Or _
                   d < ClipTop Or d >= ClipTop + Disp.HimetricToPixelY(ClipPic.Height) Then
                    A = 0
                Else
                    A = 1
                End If
            End If
            
            If bMagnify Then
                With rcPixel
                    .Left = Int(XDest + (X - CPOX) * Magnify)
                    .Top = Int(YDest + (Y - CPOY) * Magnify)
                    .Right = .Left + Magnify
                    .Bottom = .Top + Magnify
                End With
            End If
                            
            SelectClipRgn hdcFore, hrgnClip
                            
            If FinalColor >= 0 Then
                ColorToRGB FinalColor, RedTR, GreenTR, BlueTR
                ColorToRGB GetPixel(hdcFore, XDest + (X - CPOX) * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, 1, 0), YDest + (Y - CPOY) * IIf(bMagnify, Magnify, 1) + IIf(bMagnify, 1, 0)), RedTL, GreenTL, BlueTL
                If RedTL >= 0 And GreenTL >= 0 And BlueTL >= 0 Then
                    DrawColor = RGB((RedTL * Translucent + RedTR * (100 - Translucent)) / 100, _
                                    (GreenTL * Translucent + GreenTR * (100 - Translucent)) / 100, _
                                    (BlueTL * Translucent + BlueTR * (100 - Translucent)) / 100)
                Else
                    FinalColor = TransColor
                End If
            Else
                FinalColor = TransColor
            End If

            If A > 0 And FinalColor <> TransColor Then
                If bErase Then
                    If bMagnify Then
                        hbrushPixel = CreateSolidBrush(IIf(TransColor >= 0, TransColor, 0))
                    Else
                        SetPixelV hdcFore, Int(XDest + X - CPOX), Int(YDest + Y - CPOY), IIf(TransColor >= 0, TransColor, 0)
                    End If
                Else
                    If bMagnify Then
                        hbrushPixel = CreateSolidBrush(DrawColor)
                    Else
                        SetPixelV hdcFore, Int(XDest + X - CPOX), Int(YDest + Y - CPOY), DrawColor
                    End If
                End If
                If bMagnify Then
                    FillRect hdcFore, rcPixel, hbrushPixel
                    DeleteObject hbrushPixel
                End If
            End If
        Next
    Next

    SelectClipRgn hdcFore, 0
    DeleteObject hrgnClip
    DeleteObject hrgnPic

    Disp.ReleaseDC 1
    hdcFore = 0
        
    CopyBuf True, XDest + (Int(-RotWidth / 2) * IIf(bMagnify, Magnify, 1)), _
                   YDest + (Int(-RotHeight / 2) * IIf(bMagnify, Magnify, 1)), _
                   (RotWidth + 2) * IIf(bMagnify, Magnify, 1), _
                   (RotHeight + 2) * IIf(bMagnify, Magnify, 1)
        
    Exit Sub

RotateErr:
    If hDC Then Disp.ReleaseDC
    hDC = 0
    Disp.Close
    MsgBox Err.Description, vbCritical
End Sub

Sub ColorToRGB(ByVal Colr As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
Attribute ColorToRGB.VB_Description = "Split a color into its RGB components."
    R = Colr Mod 256
    G = Colr \ 256 Mod 256
    B = Colr \ 65536
End Sub

Sub ColorToRGBByte(ByVal Colr As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
Attribute ColorToRGBByte.VB_Description = "Split a color into its RGB components output as bytes."
    R = Colr Mod 256
    G = Colr \ 256 Mod 256
    B = Colr \ 65536
End Sub

Sub DrawAngle()
Attribute DrawAngle.VB_Description = "Draw the currently selected rotation angle value."
    Dim bTempDC As Boolean
    Dim rcRect As RECT
    Dim hOldFont As Long

    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If

    hOldFont = SelectObject(hDC, ButtonFont.hFont)
    SetBkColor hDC, &HC0C0C0
    
    With rcRect
        .Left = RotX - 20
        .Top = RotY - 25
        .Right = .Left + 40
        .Bottom = .Top + 25
    End With
    
    DrawText hDC, CStr(RotAngle) & "°" & Space$(8) & Chr$(0), -1, rcRect, DT_SINGLELINE
    
    SelectObject hDC, hOldFont

    CopyBuf False, rcRect.Left, rcRect.Top, 40, 25

    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub CalcRotSize(ByVal Wid As Integer, ByVal Hgt As Integer, ByRef RWid As Double, ByRef RHgt As Double)
Attribute CalcRotSize.VB_Description = "Calculate the dimensions of the specified rectangle after rotating by the currently selected angle."
    RWid = Abs(Cos(RotAngle * Pi / 180#) * Wid) + Abs(Sin(RotAngle * Pi / 180#) * Hgt)
    RHgt = Abs(Cos(RotAngle * Pi / 180#) * Hgt) + Abs(Sin(RotAngle * Pi / 180#) * Wid)
End Sub

Sub FlipClip(ByVal Ax As Integer)
Attribute FlipClip.VB_Description = "Invert the copied/clipboard image horizontally (1) or vertically (2)."
    Dim bTempDC As Boolean
    Dim Wid As Integer, Hgt As Integer
    
    If ClipPic Is Nothing Then Exit Sub
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If

    Wid = Disp.HimetricToPixelX(ClipPic.Width)
    Hgt = Disp.HimetricToPixelY(ClipPic.Height)
    
    StretchBlt hDC, ClipLeft, ClipTop, Wid, Hgt, hDC, ClipLeft + IIf(Ax = 1, Wid - 1, 0), ClipTop + IIf(Ax = 2, Hgt - 1, 0), IIf(Ax = 1, -Wid, Wid), IIf(Ax = 2, -Hgt, Hgt), SRCCOPY
    Set ClipPic = CapturePicture(hDC, ClipLeft, ClipTop, Wid, Hgt)

    CopyBuf False, ClipLeft, ClipTop, Wid, Hgt

    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub PushUndo()
Attribute PushUndo.VB_Description = "Add the current image to the undo stack."
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    CurUndo = (CurUndo + 1) Mod (UBound(UndoBuf) + 1)
    Set UndoBuf(CurUndo) = CapturePicture(hDC, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight)

    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub Undo()
Attribute Undo.VB_Description = "Handle the Undo button"
    
    If hDC Then
        Disp.Close
        MsgBox "Device context already locked"
        Exit Sub
    End If

    CurUndo = CurUndo - 1
    If CurUndo < 0 Then CurUndo = UBound(UndoBuf)
    If UndoBuf(CurUndo) Is Nothing Then
        CurUndo = (CurUndo + 1) Mod (UBound(UndoBuf) + 1)
        Exit Sub
    End If
    Disp.PaintPicture UndoBuf(CurUndo), ImageLeft, ImageTop
    Set TilePic = UndoBuf(CurUndo)
    DrawGrid
    PressThrough True, True
        
End Sub

Sub Redo()
Attribute Redo.VB_Description = "Process the Redo button."
    
    If hDC Then
        Disp.Close
        MsgBox "Device context already locked"
        Exit Sub
    End If
        
    CurUndo = (CurUndo + 1) Mod (UBound(UndoBuf) + 1)
    If UndoBuf(CurUndo) Is Nothing Then
        CurUndo = CurUndo - 1
        If CurUndo < 0 Then CurUndo = UBound(UndoBuf)
        Exit Sub
    End If
    Disp.PaintPicture UndoBuf(CurUndo), ImageLeft, ImageTop
    Set TilePic = UndoBuf(CurUndo)
    DrawGrid
    PressThrough True, True
    
End Sub

Sub MakeBlankTile(ByVal TileWidth, ByVal TileHeight)
Attribute MakeBlankTile.VB_Description = "Load the grid/editor with an empty tile."
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    Set TilePic = CapturePicture(hDC, -TileWidth, -TileHeight, TileWidth, TileHeight)
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub PicShift(ByVal DX As Integer, ByVal DY As Integer)
Attribute PicShift.VB_Description = "Shift the pixels in the editor image."
    Dim bTempDC As Boolean
    
    If TilePic Is Nothing Then Exit Sub
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    BitBlt hDC, ImageLeft + IIf(DX > 0, DX, 0), ImageTop + IIf(DY > 0, DY, 0), Tileset.TileWidth - Abs(DX), Tileset.TileHeight - Abs(DY), hDC, ImageLeft - IIf(DX < 0, DX, 0), ImageTop - IIf(DY < 0, DY, 0), SRCCOPY
    Set TilePic = CapturePicture(hDC, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight)
    PressThrough True, False

    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If

End Sub

Sub DrawTileset()
Attribute DrawTileset.VB_Description = "Draw the tile selection ""window"" onto the display."
    Dim SourceXOff As Integer
    Dim SourceYOff As Integer
    Dim CellXOff As Integer
    Dim CellYOff As Integer
    Dim rcCell As RECT
    
    If hDC Then
        Disp.Close
        MsgBox "Device context already locked in DrawTileSet"
        Exit Sub
    End If
    
    hDC = Disp.GetDC
    With rcCell
        .Left = TSViewLeft - 2
        .Top = TSViewTop - 2
        .Right = .Left + TSViewWidth + 4
        .Bottom = .Top + TSViewHeight + 4
    End With
    DrawFrameControl hDC, rcCell, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED
    Disp.ReleaseDC
    hDC = 0
        
    CellXOff = Tileset.TileWidth * (CurCell Mod Tileset.TileColumns)
    CellYOff = Tileset.TileHeight * (CurCell \ Tileset.TileColumns)
    
    SourceXOff = CellXOff - (TSViewWidth - Tileset.TileWidth) \ 2
    SourceYOff = CellYOff - (TSViewHeight - Tileset.TileHeight) \ 2
    
    Disp.CopyRect Tileset, 0, 0, Tileset.TileWidth * Tileset.TileColumns, Tileset.TileHeight * TileRows, TSViewLeft - SourceXOff, TSViewTop - SourceYOff, TSViewLeft, TSViewTop, TSViewWidth, TSViewHeight, False
    
    With rcCell
        .Left = TSViewLeft + (TSViewWidth - Tileset.TileWidth) \ 2
        .Top = TSViewTop + (TSViewHeight - Tileset.TileHeight) \ 2
        .Right = .Left + Tileset.TileWidth
        .Bottom = .Top + Tileset.TileHeight
    End With
    
    hDC = Disp.GetDC
    DrawFocusRect hDC, rcCell
    CopyBuf False, TSViewLeft, TSViewTop, TSViewWidth, TSViewHeight
    Disp.ReleaseDC
    hDC = 0
    
End Sub

Sub TileSetPress(ByVal X As Integer, ByVal Y As Integer)
Attribute TileSetPress.VB_Description = "Handle a mouse click on the tileset window (select a tile)."
    Dim XOff As Integer, YOff As Integer
    Dim NewCell As Integer
    
    If X <= TSViewLeft Or Y <= TSViewTop Or X >= TSViewLeft + TSViewWidth Or Y > TSViewTop + TSViewHeight Then Exit Sub
    XOff = X - (TSViewLeft + (TSViewWidth - Tileset.TileWidth) \ 2)
    YOff = Y - (TSViewTop + (TSViewHeight - Tileset.TileHeight) \ 2)
    NewCell = CurCell + Int(XOff / Tileset.TileWidth) + Int(YOff / Tileset.TileHeight) * Tileset.TileColumns
    If NewCell >= 0 And NewCell < Tileset.TileColumns * TileRows Then
        CurCell = NewCell
        DrawTileset
    End If
    
End Sub

Sub CopyBuf(ByVal ForeToBack As Boolean, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0, Optional ByVal Width As Integer = 640, Optional ByVal Height As Integer = 480)
Attribute CopyBuf.VB_Description = "Copy a portion of the background buffer onto the foreground buffer of the display, or vice versa."
    Dim bTempDC As Boolean
    Dim hdcFore As Long
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    hdcFore = Disp.GetDC(1)
    If ForeToBack Then
        BitBlt hDC, X, Y, Width, Height, hdcFore, X, Y, SRCCOPY
    Else
        BitBlt hdcFore, X, Y, Width, Height, hDC, X, Y, SRCCOPY
    End If
    Disp.ReleaseDC 1
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
End Sub

Sub PressThrough(Optional ByVal Small As Boolean = True, Optional ByVal Mag As Boolean = True)
Attribute PressThrough.VB_Description = "Copy the grid and/or un-magnified image from the back buffer to the front buffer."
    If Mag Then CopyBuf False, 0, 0, Tileset.TileWidth * Magnify + 1, Tileset.TileHeight * Magnify + 1
    If Small Then CopyBuf False, ImageLeft, ImageTop, Tileset.TileWidth, Tileset.TileHeight
End Sub

Private Function Mix9Colors(ByVal TLColor As Long, ByVal TColor As Long, ByVal TRColor As Long, ByVal LColor As Long, ByVal CenterColor As Long, ByVal RColor As Long, ByVal BLColor As Long, ByVal BColor As Long, ByVal BRColor As Long) As Long
    Dim Red(8) As Integer, Green(8) As Integer, Blue(8) As Integer
    Dim ResultRed As Integer, ResultGreen As Integer, ResultBlue As Integer
    Dim I As Integer
    Dim Weight As Integer
    
    ColorToRGB CenterColor, Red(0), Green(0), Blue(0)
    ColorToRGB TColor, Red(1), Green(1), Blue(1)
    ColorToRGB LColor, Red(2), Green(2), Blue(2)
    ColorToRGB RColor, Red(3), Green(3), Blue(3)
    ColorToRGB BColor, Red(4), Green(4), Blue(4)
    ColorToRGB TLColor, Red(5), Green(5), Blue(5)
    ColorToRGB TRColor, Red(6), Green(6), Blue(6)
    ColorToRGB BLColor, Red(7), Green(7), Blue(7)
    ColorToRGB BRColor, Red(8), Green(8), Blue(8)
    
    For I = 0 To 8
        Select Case I
        Case 0: Weight = 8
        Case 1 To 4: Weight = 2
        Case Else: Weight = 1
        End Select
        ResultRed = ResultRed + Red(I) * Weight
        ResultGreen = ResultGreen + Green(I) * Weight
        ResultBlue = ResultBlue + Blue(I) * Weight
    Next I
    
    Mix9Colors = RGB(ResultRed \ 20, ResultGreen \ 20, ResultBlue \ 20)
    
End Function

Public Property Get TileSetBitmap() As StdPicture
Attribute TileSetBitmap.VB_Description = "Return the image of the tileset currently being edited."
    Set TileSetBitmap = Tiles
End Property

Sub Palette16Bit()
Attribute Palette16Bit.VB_Description = "Force the entries in the palette to be compatible with 16-bit color (so floodfill doesn't dither)."
    Dim R As Integer, G As Integer, B As Integer
    Dim I As Integer
    
    For I = LBound(PalEntries) To UBound(PalEntries)
        ColorToRGB PalEntries(I), R, G, B
        R = (R \ 8) * 8
        G = (G \ 8) * 8
        B = (B \ 8) * 8
        PalEntries(I) = RGB(R, G, B)
    Next I
    DrawPalette
    
End Sub

Sub MakeGridBuf()
Attribute MakeGridBuf.VB_Description = "Create a buffer to temporarily store the image of the grid."
    Dim bTempDC As Boolean
    
    If hdcGridBuf Or hbmpGridBuf Then
        Disp.Close
        Err.Raise vbObjectError, , "Grid buffer already created"
    End If
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    hdcGridBuf = CreateCompatibleDC(hDC)
    hbmpGridBuf = CreateCompatibleBitmap(hDC, Tileset.TileWidth * Magnify, Tileset.TileHeight * Magnify)
    hbmpGridNull = SelectObject(hdcGridBuf, hbmpGridBuf)
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Sub StoreGrid()
Attribute StoreGrid.VB_Description = "Store the magnified (grid) image into the grid buffer."
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    BitBlt hdcGridBuf, 0, 0, Tileset.TileWidth * Magnify, Tileset.TileHeight * Magnify, hDC, GridLeft, GridTop, SRCCOPY
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
End Sub

Sub RestoreGrid()
Attribute RestoreGrid.VB_Description = "Restore the stored image of the grid area."
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    BitBlt hDC, GridLeft, GridTop, Tileset.TileWidth * Magnify, Tileset.TileHeight * Magnify, hdcGridBuf, 0, 0, SRCCOPY
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
End Sub

Sub DeleteGridBuf()
Attribute DeleteGridBuf.VB_Description = "Delete the buffer for temporary storage of the grid's image."
    If hdcGridBuf Then
        SelectObject hdcGridBuf, hbmpGridNull
        DeleteObject hbmpGridBuf
        DeleteDC hdcGridBuf
        hbmpGridBuf = 0
        hdcGridBuf = 0
    End If
End Sub

Sub SetClipping(ByVal Area As Integer)
Attribute SetClipping.VB_Description = "0=No clipping, 1=Clip to grid, 2=Clip to un-magnified image."
    Dim ClipRgn As Long
    Dim AddRgn As Long
    Dim bTempDC As Boolean
    
    If hDC = 0 Then
        hDC = Disp.GetDC
        bTempDC = True
    Else
        bTempDC = False
    End If
    
    Select Case Area
    Case 0 ' Off
        SelectClipRgn hDC, 0
    Case 1 ' Grid
        ClipRgn = CreateRectRgn(GridLeft, GridTop, GridLeft + Tileset.TileWidth * Magnify, GridTop + Tileset.TileHeight * Magnify)
        SelectClipRgn hDC, ClipRgn
        DeleteObject ClipRgn
    Case 2 ' Image
        ClipRgn = CreateRectRgn(ImageLeft, ImageTop, ImageLeft + Tileset.TileWidth, ImageTop + Tileset.TileHeight)
        SelectClipRgn hDC, ClipRgn
        DeleteObject ClipRgn
    End Select
    
    If bTempDC Then
        Disp.ReleaseDC
        hDC = 0
    End If
    
End Sub

Property Set TilePicture(P As StdPicture)
Attribute TilePicture.VB_Description = "Set a new tile image into the grid/editor."
    Set TilePic = P
    DrawAll
    'Disp.PaintPicture TilePic, ImageLeft, ImageTop
    'DrawGrid
    'PressThrough
End Property

Public Property Get PalEntry(ByVal Index As Integer) As Long
Attribute PalEntry.VB_Description = "Set/get the color for a particular entry in the palette."
    PalEntry = PalEntries(Index)
End Property

Public Property Let PalEntry(ByVal Index As Integer, ByVal NewVal As Long)
    PalEntries(Index) = NewVal
End Property
