' This is a generic tileset editing template script
' It connects to the tileset editor whenever it comes up.
' To customize the script to your own purpose, edit
' the second section (script number 1) in this file.

' ======== INITIAL STARTUP SCRIPT (Number 0) =========
Sub Project_OnEditTileset(T)
   Set HostObj.TempStorage = T
   HostObj.StartScript=1
End Sub

HostObj.SinkObjectEvents ProjectObj, "Project"

HostObj.ConnectEventsNow()

#Split == MAP EDITING SCRIPT (Number 1) ============

Dim SMX, SMY ' Mouse X and Mouse Y screen corrdinates
Dim DispMsg
Dim TsEd

Set TsEd = HostObj.TempStorage
HostObj.TempStorage = Empty

' Delete code for events that you don't need

Sub TileEd_OnEditInit()
   ' Insert Your TileEdit Initialization code here
   TsEd.Disp.Forecolor = 0
   For Rpt = 1 to 2
      TsEd.Disp.DrawText "Tile editing script is connected" & DispMsg, 400, 465
      TsEd.Disp.Flip
   Next
End Sub

Sub TileEd_OnKeyPress(KeyAscii)
   ' Insert your special keypress handling here
   ' (This sample code re-arranges the tiles of the free, but not freely distributable
   '  SpriteLib graphics library copyrighted by Ari Feldman.  It applies to the Blocks1
   '  and Blocks2 bitmaps. (http://www.arifeldman.com/))
   'For Y=0 to 10
   '   For X = 0 to 17
   '      Set TmpPic = HostObj.ExtractPic(TsEd.TileSetBitmap, X*34+2,Y*34+2,32,32,False)
   '      HostObj.PasteTileToPic TsEd.TileSetBitmap, TmpPic, X*32, Y*32
   '   Next
   'Next
   'TsEd.Disp.Tilesets(0).PaintPicture TsEd.TileSetBitmap,0,0
   'TsEd.DrawAll
End Sub

Sub TileEd_OnMouseMove(Button, Shift, X, Y)
   ' Add special mousemove code here
   SMX = X
   SMY = Y
End Sub

Sub TileEd_OnMouseDown(Button, Shift, X, Y)
   ' Add special mousedown code here
End Sub

Sub TileEd_OnMouseUp(Button, Shift, X, Y)
   ' Add special mouseup code here
End Sub

Sub TileEd_OnEditComplete()
   Set TsEd = Nothing
   HostObj.StartScript=0
End Sub

HostObj.SinkObjectEvents HostObj.AsObject(TsEd), "TileEd"
HostObj.ConnectEventsNow()