' This is a generic map editing template script
' It connects to the map editor whenever it comes up.
' To customize the script to your own purpose, edit
' the second section (script number 1) in this file.

' ======== INITIAL STARTUP SCRIPT (Number 0) =========
Sub Project_OnEditMap(M)
   Set HostObj.TempStorage = M
   HostObj.StartScript=1
End Sub

HostObj.SinkObjectEvents ProjectObj, "Project"

HostObj.ConnectEventsNow()

#Split == MAP EDITING SCRIPT (Number 1) ============

Dim SMX, SMY ' Mouse X and Mouse Y screen corrdinates
Dim MpEd

Set MpEd = HostObj.TempStorage
HostObj.TempStorage = Empty

' Delete code for events that you don't need

Sub MapEd_OnEditInit()
   ' Insert Your Map Initialization code here
   MpEd.DisplayMessage = "Map editing script is connected"
End Sub

Sub MapEd_OnKeyPress(KeyAscii)
   ' Add special keyboard handling here
   If KeyAscii <> 13 Then MpEd.DisplayMessage = "Script trapped key ASCII:" & KeyAscii
End Sub

Sub MapEd_OnAfterDraw()
   ' Add special map drawing code here
   MpEd.Disp.DrawText "[S" & SMX & "," & SMY & "]", 500, 0
End Sub

Sub MapEd_OnMouseDown(Button, Shift, X, Y)
   ' Add special mousedown code here
   MpEd.DisplayMessage = "Script trapped MouseDown at " & CStr(X) & ", " & CStr(Y)
End Sub

Sub MapEd_OnMouseUp(Button, Shift, X, Y)
   ' Add special mouseup code here
   MpEd.DisplayMessage = ""
End Sub

Sub MapEd_OnMouseMove(Button, Shift, X, Y)
   SMX = X
   SMY = Y
End Sub

Sub MapEd_OnEditComplete()
   HostObj.StartScript=0
End Sub

HostObj.SinkObjectEvents HostObj.AsObject(MpEd), "MapEd"
HostObj.ConnectEventsNow()