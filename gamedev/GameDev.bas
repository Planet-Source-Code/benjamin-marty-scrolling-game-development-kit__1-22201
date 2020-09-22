Attribute VB_Name = "Globals"
Attribute VB_Description = "Global utility functions, variables (Project) and API declarations here."
'======================================================================
'
' Project: GameDev - Scrolling Game Development Kit
'
' Developed By Benjamin Marty
' Copyright © 2000,2001 Benjamin Marty
' Distibuted under the GNU General Public License
'    - see http://www.fsf.org/copyleft/gpl.html
'
' File: GameDev.bas - Global Code and Declarations Module
'                     and Main Starting Point
'
'======================================================================

Option Explicit

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type IID
    X As Long
    s1 As Integer
    s2 As Integer
    C(0 To 7) As Byte
End Type

Public Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hBitmap As Long
    hpal As Long
End Type

Public Type Size
        cx As Long
        cy As Long
End Type

Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const OBJ_PAL = 5
Public Const PICTYPE_BITMAP = 1
Public Const S_OK = &H0
Public Const WHITE_PEN = 6
Public Const WHITE_BRUSH = 0
Public Const BLACK_BRUSH = 4
Public Const DEFAULT_GUI_FONT = 17
Public Const DFC_BUTTON = 4
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_PUSHED = &H200
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const DT_WORDBREAK = &H10
Public Const DT_CALCRECT = &H400
Public Const Transparent = 1
Public Const LTGRAY_BRUSH = 1
Public Const DKGRAY_BRUSH = 3
Public Const FLOODFILLSURFACE = 1
Public Const NULL_BRUSH = 5
Public Const GRAY_BRUSH = 2
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const RGN_OR = 2
Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const JOYERR_NOERROR = 0  '  no error
Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10                             '  axis 5
Public Const JOY_RETURNV = &H20                             '  axis 6
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000

Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary

Public Const Pi As Double = 3.14159265358979

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Attribute SelectObject.VB_Description = "Win32 API"
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Attribute MoveToEx.VB_Description = "Win32 API"
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Attribute LineTo.VB_Description = "Win32 API"
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Attribute Rectangle.VB_Description = "Win32 API"
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Attribute CreateCompatibleBitmap.VB_Description = "Win32 API"
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Attribute CreateCompatibleDC.VB_Description = "Win32 API"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Attribute BitBlt.VB_Description = "Win32 API"
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Attribute DeleteDC.VB_Description = "Win32 API"
Public Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As IID, ByVal fOwn As Boolean, ByRef ppvObj As StdPicture) As Long
Attribute OleCreatePictureIndirect.VB_Description = "Win32 API"
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Attribute DeleteObject.VB_Description = "Win32 API"
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Attribute GetCurrentObject.VB_Description = "Win32 API"
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Attribute GetDC.VB_Description = "Win32 API"
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Attribute ReleaseDC.VB_Description = "Win32 API"
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Attribute CreateSolidBrush.VB_Description = "Win32 API"
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Attribute GetPixel.VB_Description = "Win32 API"
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Attribute FillRect.VB_Description = "Win32 API"
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Attribute GetStockObject.VB_Description = "Win32 API"
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Attribute DrawFrameControl.VB_Description = "Win32 API"
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Attribute GetTextExtentPoint32.VB_Description = "Win32 API"
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Attribute DrawText.VB_Description = "Win32 API"
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Attribute SetBkMode.VB_Description = "Win32 API"
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Attribute SetPixel.VB_Description = "Win32 API"
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Attribute CreatePen.VB_Description = "Win32 API"
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Attribute ExtFloodFill.VB_Description = "Win32 API"
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Attribute Ellipse.VB_Description = "Win32 API"
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Attribute SetBkColor.VB_Description = "Win32 API"
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Attribute DrawFocusRect.VB_Description = "Win32 API"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Attribute SetPixelV.VB_Description = "Win32 API"
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Attribute StretchBlt.VB_Description = "Win32 API"
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Attribute SetCursor.VB_Description = "Win32 API"
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByRef lpFilePart As Long) As Long
Attribute GetFullPathName.VB_Description = "Win32 API"
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Attribute CreateRectRgn.VB_Description = "Win32 API"
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Attribute SelectClipRgn.VB_Description = "Win32 API"
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Attribute CombineRgn.VB_Description = "Win32 API"
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Attribute GetDIBits.VB_Description = "Win32 API"
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Attribute CreateDIBSection.VB_Description = "Win32 API"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal length As Long)
Attribute CopyMemory.VB_Description = "Win32 API"
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Attribute SetDIBits.VB_Description = "Win32 API"
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Attribute RegisterClipboardFormat.VB_Description = "Win32 API"
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Attribute SetTextColor.VB_Description = "Win32 API"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Attribute GetKeyState.VB_Description = "Win32 API"
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Attribute IntersectRect.VB_Description = "Win32 API"
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Attribute joyGetPosEx.VB_Description = "Win32 API"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Attribute Sleep.VB_Description = "Win32 API"
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetDwValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function VB6fCreateShellLink Lib "VB6STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Public Declare Function VB5fCreateShellLink Lib "VB5STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Public Prj As GameProject
Attribute Prj.VB_VarDescription = "The global instance of the game project."
Public GameHost As ScriptHost
Attribute GameHost.VB_VarDescription = "Global instance of the script hosting component that handles script for this instance of GameDev."
Public CurDisp As BMDXDisplay
Attribute CurDisp.VB_VarDescription = "Global reference to the currently open BMDXDisplay.  Nothing if not open."
Public LoadVersion As Integer
Attribute LoadVersion.VB_VarDescription = "The version number of the file currently being loaded."

Public Sub RegGDP()
    Dim hkcr As Long
    Dim hGDPKey As Long
    Dim hGDPFileKey As Long
    Dim hGDPShellKey As Long
    Dim Disposition As Long
    Dim strVal As String
    Dim lResult As Long
    hkcr = &H80000000 ' HKEY_CLASSES_ROOT
    
    If GetSetting("GameDev", "Switch", "AskRegister") = "No" And (Abs(GetKeyState(vbKeyShift)) <= 1) Then Exit Sub
    lResult = RegOpenKeyEx(hkcr, ".gdp", 0, KEY_QUERY_VALUE, hGDPKey)
    If lResult <> ERROR_SUCCESS Then
        frmSplash.Visible = False
        Select Case MsgBox("The extension .GDP has not been registered on this system.  Would you like to associate "".GDP"" with GameDev now? (Selecting ""No"" won't ask this question again.)", vbQuestion + vbYesNoCancel, "GameDev")
        Case vbNo
            SaveSetting "GameDev", "Switch", "AskRegister", "No"
            frmSplash.Visible = True
            Exit Sub
        Case vbCancel
            frmSplash.Visible = True
            Exit Sub
        End Select
        frmSplash.Visible = True
    Else
        RegCloseKey hGDPKey
    End If
    
    If ERROR_SUCCESS <> RegCreateKeyEx(hkcr, ".gdp", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPKey, Disposition) Then
        MsgBox "Failed to register GDP file type."
        Exit Sub
    End If
    
    If ERROR_SUCCESS <> RegCreateKeyEx(hkcr, "gdpfile", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPFileKey, Disposition) Then
        GoTo CloseKeys
        Exit Sub
    End If
    
    lResult = RegSetValueEx(hGDPKey, "", 0, REG_SZ, "gdpfile", 8)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys
    
    lResult = RegSetValueEx(hGDPFileKey, "", 0, REG_SZ, "GameDev Project", 16)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys
    
    Disposition = 0
    lResult = RegSetDwValueEx(hGDPFileKey, "EditFlags", 0, REG_BINARY, Disposition, 4)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys
    
    RegCloseKey hGDPKey
    lResult = RegCreateKeyEx(hGDPFileKey, "DefaultIcon", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPKey, Disposition)
    If ERROR_SUCCESS <> lResult Then GoTo CloseKeys
    
    strVal = App.Path & "\" & App.EXEName & ".exe,0"
    lResult = RegSetValueEx(hGDPKey, "", 0, REG_SZ, strVal, Len(strVal) + 1)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys

    lResult = RegCreateKeyEx(hGDPFileKey, "Shell", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPShellKey, Disposition)
    If ERROR_SUCCESS <> lResult Then GoTo CloseKeys

    lResult = RegSetValueEx(hGDPShellKey, "", 0, REG_SZ, "Edit", 5)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys
    
    RegCloseKey hGDPKey
    lResult = RegCreateKeyEx(hGDPShellKey, "Edit\command", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPKey, Disposition)
    If ERROR_SUCCESS <> lResult Then GoTo CloseKeys

    strVal = """" & App.Path & "\" & App.EXEName & ".exe"" ""%1"""
    lResult = RegSetValueEx(hGDPKey, "", 0, REG_SZ, strVal, Len(strVal) + 1)
    If lResult <> ERROR_SUCCESS Then GoTo CloseKeys
    
    RegCloseKey hGDPKey
    lResult = RegCreateKeyEx(hGDPShellKey, "Play\command", 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_QUERY_VALUE Or KEY_SET_VALUE, 0, hGDPKey, Disposition)
    If ERROR_SUCCESS <> lResult Then GoTo CloseKeys
    
    strVal = """" & App.Path & "\" & App.EXEName & ".exe"" ""%1"" /p"
    lResult = RegSetValueEx(hGDPKey, "", 0, REG_SZ, strVal, Len(strVal) + 1)

CloseKeys:
    RegCloseKey hGDPKey
    RegCloseKey hGDPShellKey
    RegCloseKey hGDPFileKey
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Failed to register GDP file type."
    End If
End Sub

Public Function CapturePicture(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As IPictureDisp
Attribute CapturePicture.VB_Description = "Capture a piece of the display (or any device context) into an OLE Picture object."
    Dim picGuid As IID
    Dim picDesc As PICTDESC
    Dim hdcMem As Long
    Dim hBmp As Long
    Dim hOldBmp As Long
    Dim Pic As IPictureDisp
    Dim rcBitmap As RECT
    
    ' IID_IPictureDisp
    picGuid.X = &H7BF80981
    picGuid.s1 = &HBF32
    picGuid.s2 = &H101A
    picGuid.C(0) = &H8B
    picGuid.C(1) = &HBB
    picGuid.C(2) = &H0
    picGuid.C(3) = &HAA
    picGuid.C(4) = &H0
    picGuid.C(5) = &H30
    picGuid.C(6) = &HC
    picGuid.C(7) = &HAB
    
    picDesc.cbSizeOfStruct = Len(picDesc)
    hdcMem = CreateCompatibleDC(hDC)
    If hdcMem = 0 Then
        Exit Function
    End If
    hBmp = CreateCompatibleBitmap(hDC, Width, Height)
    If hBmp = 0 Then
        DeleteDC hdcMem
        Exit Function
    End If
    hOldBmp = SelectObject(hdcMem, hBmp)
    If Left >= 0 And Top >= 0 Then
        If BitBlt(hdcMem, 0, 0, Width, Height, hDC, Left, Top, SRCCOPY) = 0 Then
            SelectObject hdcMem, hOldBmp
            DeleteDC hdcMem
            DeleteObject hBmp
            Exit Function
        End If
    Else
        With rcBitmap
           .Left = 0
           .Top = 0
           .Right = Width
           .Bottom = Height
        End With
        FillRect hdcMem, rcBitmap, GetStockObject(BLACK_BRUSH)
    End If
    SelectObject hdcMem, hOldBmp
    picDesc.hBitmap = hBmp
    picDesc.hpal = GetCurrentObject(hDC, OBJ_PAL)
    picDesc.picType = PICTYPE_BITMAP
    If OleCreatePictureIndirect(picDesc, picGuid, True, Pic) <> S_OK Then
        DeleteDC hdcMem
        DeleteObject hBmp
        Exit Function
    End If
    Set CapturePicture = Pic
    Set Pic = Nothing
    DeleteDC hdcMem
End Function

Public Function GetRelativePath(ByVal BasePath As String, ByVal RelPath As String) As String
Attribute GetRelativePath.VB_Description = "Given a path and filename (RelPath), calculate it's path relative to another base path (BasePath)."
    Dim I As Long
    Dim L As Long
    Dim Base As String
    Dim Rel As String
    
    Base = Space$(256)
    L = GetFullPathName(BasePath & Chr$(0), Len(Base), Base, I)
    Base = Left$(Base, L)
    If L > 0 Then
        Do While L > 1 And Mid$(Base, L, 1) <> "\"
            L = L - 1
        Loop
        If InStr(Mid(Base, L + 1), ".") Then
            Base = Left$(Base, L)
        Else
            If Right$(Base, 1) <> "\" Then Base = Base & "\"
        End If
    End If
    Rel = Space$(256)
    L = GetFullPathName(RelPath & Chr$(0), Len(Rel), Rel, I)
    Rel = Left$(Rel, L)
    
    If UCase$(Left$(Base, 2)) <> UCase$(Left$(Rel, 2)) Then
        GetRelativePath = Rel
        Exit Function
    End If
    
    I = 3
    Do While UCase$(Mid$(Base, I, 1)) = UCase$(Mid$(Rel, I, 1)) And I <= Len(Base)
        I = I + 1
    Loop
    
    Do While I > 2 And Mid$(Base, I - 1, 1) <> "\"
        I = I - 1
    Loop
    
    L = I
    Do While L < Len(Base)
        L = L + 1
        If Mid$(Base, L, 1) = "\" Then GetRelativePath = GetRelativePath & "..\"
    Loop
    
    GetRelativePath = GetRelativePath & Mid$(Rel, I)
End Function

Function ParseArg(ByRef Params As String) As String
Attribute ParseArg.VB_Description = "Parse a command line argument to GameDev."
    Dim SepPos As Integer
    Dim ArgStart As Integer

    If Left$(Params, 1) = """" Then
        SepPos = InStr(2, Params, """", vbTextCompare)
        ArgStart = 2
    Else
        SepPos = InStr(Params, " ")
        ArgStart = 1
    End If

    If SepPos = 0 Then
        ParseArg = Mid$(Params, ArgStart)
        Params = ""
    Else
        ParseArg = Mid$(Params, ArgStart, SepPos - ArgStart)
        Params = LTrim$(Mid$(Params, SepPos + 1))
    End If

End Function

Sub Main()
Attribute Main.VB_Description = "This is where execution of GameDev starts."
    Dim StartupScript As String
    Dim ScriptName As String
    Dim bPlayOnly As Boolean
    Dim CmdLine As String
    Dim Arg As String
    Dim TmpArgs As String
    Dim ScreenDepth As Integer

    If Len(Dir$(App.Path & "\Help\GameDev.chm")) > 0 Then
        App.HelpFile = App.Path & "\Help\GameDev.chm"
    Else
        App.HelpFile = App.Path & "\GameDev.chm"
    End If
    frmSplash.Show
    frmSplash.Refresh
    Set Prj = New GameProject
    Prj.bSplashShowing = True
    CmdLine = Command$
    ScreenDepth = Val(GetSetting("GameDev", "Options", "ScreenDepth", "16"))
    If ScreenDepth <> 16 And ScreenDepth <> 24 And ScreenDepth <> 32 Then ScreenDepth = 16
    Do While Len(CmdLine)
        Arg = ParseArg(CmdLine)
        If Left$(Arg, 1) = "/" Or Left$(Arg, 1) = "-" Then
            Select Case UCase$(Mid$(Arg, 2, 1))
            Case "P", "E"
                If UCase$(Mid$(Arg, 2, 1)) = "P" Then bPlayOnly = True
                If Len(CmdLine) = 0 Then
                    ScriptName = App.Path & "\GameDev.vbs"
                Else
                    TmpArgs = CmdLine
                    Arg = ParseArg(TmpArgs)
                    If Left$(Arg, 1) <> "/" And Left$(Arg, 1) <> "-" Then
                        ScriptName = ParseArg(CmdLine)
                    Else
                        ScriptName = App.Path & "\GameDev.vbs"
                    End If
                End If
            Case "D"
                ScreenDepth = Val(ParseArg(CmdLine))
                If ScreenDepth <> 16 And ScreenDepth <> 24 And ScreenDepth <> 32 Then
                    MsgBox "Invalid screen depth """ & ScreenDepth & """.", vbExclamation
                    Unload frmSplash
                    Exit Sub
                End If
            Case "?"
                frmSplash.tmrUnload.Enabled = False
                MsgBox "GameDev [<ProjectFile.gdp>] [/p [<Script.vbs>] [/d <screen depth>] | /e [<Script.vbs>]]" & vbCrLf & _
                       "/p - Play script and exit." & vbCrLf & "/e - Play script and load editor." & vbCrLf & _
                       "/d - Specify color bit depth for playing (16/24/32)", vbInformation
                Unload frmSplash
            Case Else
                MsgBox "Invalid switch """ & UCase$(Mid$(Arg, 2, 1)) & """.", vbExclamation
                Unload frmSplash
                Exit Sub
            End Select
        Else
            If Len(Dir$(Arg)) > 0 Then
                Prj.Load Arg
            Else
                MsgBox "Cannot find " & Arg, vbExclamation
                Unload frmSplash
                Exit Sub
            End If
        End If
    Loop
    
    If (Len(ScriptName) > 0) And (Len(Dir$(ScriptName)) > 0) Then
        Open ScriptName For Input As #1
            StartupScript = Input$(LOF(1), 1)
        Close
        Set GameHost = New ScriptHost
        GameHost.InitScript StartupScript
        Sleep 2000
        Unload frmSplash
        GameHost.RunStartScript
        GameHost.CheckForError
    ElseIf bPlayOnly Then
        Sleep 2000
        Unload frmSplash
        Prj.GamePlayer.Play ScreenDepth
    End If
    ' Show UI only when EXE is executed directly by user
    If App.StartMode = 0 And bPlayOnly = False Then
        RegGDP
        frmProject.Show
        frmProject.LoadTree
    End If
End Sub

Public Sub GeneralDeserialize(ByRef Data As String, ParamArray Parms() As Variant)
Attribute GeneralDeserialize.VB_Description = "Generically read a series of values (as formatted in a GDP project file) into a series of variables based on a series of names."
    Dim DatStart As Integer, DatEnd As Integer
    Dim Key As String
    Dim Value As String
    Dim FoundCount As Integer
    Dim KeyIdx As Integer
    Dim ItemCount As Integer

    ItemCount = (UBound(Parms) - LBound(Parms) + 1) / 2
    DatStart = 1
    Do While Mid$(Data, DatStart) <= " " And DatStart < Len(Data)
        DatStart = DatStart + 1
    Loop
    For FoundCount = 1 To ItemCount
        DatEnd = InStr(DatStart, Data, "=", vbTextCompare)
        If DatEnd <= 0 Then Err.Raise vbObjectError, , "Error deserializing data"
        Key = Trim$(Mid$(Data, DatStart, DatEnd - DatStart))
        DatStart = DatEnd + 1
        DatEnd = InStr(DatStart, Data, vbCrLf, vbTextCompare)
        If DatEnd <= 0 Then DatEnd = Len(Data) + 1
        Value = Trim$(Mid$(Data, DatStart, DatEnd - DatStart))
        For KeyIdx = 0 To ItemCount - 1
            If LCase$(Key) = LCase$(Parms(KeyIdx)) Then
                Select Case VarType(Parms(KeyIdx + ItemCount))
                Case vbString, vbEmpty
                    Parms(KeyIdx + ItemCount) = Value
                Case vbBoolean
                    Parms(KeyIdx + ItemCount) = CBool(Value)
                Case Else
                    Parms(KeyIdx + ItemCount) = Val(Value)
                End Select
                Exit For
            End If
        Next
        If KeyIdx >= ItemCount Then
            Err.Raise vbObjectError, , "Error parsing data; could not recognize """ & Key & """"
        End If
        DatStart = DatEnd + 2
    Next
    Data = Mid$(Data, DatStart)
End Sub

Public Function GeneralSerialize(ParamArray Parms() As Variant) As String
Attribute GeneralSerialize.VB_Description = "Generically store a series of variables into a string formatted like GDP project file based on a series of names."
    Dim I As Integer
    Dim ItemCount As Integer
    
    ItemCount = (UBound(Parms) - LBound(Parms) + 1) / 2
    GeneralSerialize = ""
    For I = 0 To ItemCount - 1
        GeneralSerialize = GeneralSerialize & Parms(I) & "=" & CStr(Parms(I + ItemCount)) & vbCrLf
    Next I
End Function

Public Function PathFromFile(ByVal FilePath As String) As String
Attribute PathFromFile.VB_Description = "Given a path and file name, return only the path."
    Dim Buf As String
    Dim L As Long
    Dim I As Long
    
    Buf = Space$(256)
    
    L = GetFullPathName(FilePath & Chr$(0), Len(Buf), Buf, I)
    Buf = Left$(Buf, L)
    Do While Mid$(Buf, L, 1) <> "\" And L > 1
        L = L - 1
    Loop
    PathFromFile = Left$(Buf, L - 1)
End Function

Public Sub PasteTileToPicture(ByVal Pic As StdPicture, ByVal Tile As StdPicture, ByVal X As Integer, ByVal Y As Integer)
Attribute PasteTileToPicture.VB_Description = "Paste one (""Tile"") OLE Picture object directly into another (""Pic"")."
    Dim hDC As Long
    Dim bmiPic As BITMAPINFO
    Dim bmiTile As BITMAPINFO
    Dim DIBits() As RGBQUAD
    Dim DIBitsPic() As RGBQUAD
    Dim YIndex As Long
    Dim XIndex As Long
    Dim OffTile As Long
    Dim OffPic As Long

    hDC = GetDC(0)
    If hDC = 0 Then
        Exit Sub
    End If
    
    With bmiPic.bmiHeader
        .biSize = Len(bmiPic.bmiHeader)
        .biPlanes = 1
    End With
    With bmiTile.bmiHeader
        .biSize = Len(bmiTile.bmiHeader)
        .biPlanes = 1
    End With
    
    If 0 = GetDIBits(hDC, Pic.handle, 0, 0, ByVal 0&, bmiPic, DIB_RGB_COLORS) Then
        ReleaseDC 0, hDC
        Exit Sub
    End If
        
    With bmiPic.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    If 0 = GetDIBits(hDC, Tile.handle, 0, 0, ByVal 0&, bmiTile, DIB_RGB_COLORS) Then
        ReleaseDC 0, hDC
        Exit Sub
    End If
    
    With bmiTile.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    If Abs(bmiTile.bmiHeader.biHeight) + Y > Abs(bmiPic.bmiHeader.biHeight) Or Y < 0 Then
        ReleaseDC 0, hDC
        Exit Sub
    End If
    
    If bmiTile.bmiHeader.biWidth + X > bmiPic.bmiHeader.biWidth Or X < 0 Then
        ReleaseDC 0, hDC
        Exit Sub
    End If
    
    ReDim DIBits(0 To bmiTile.bmiHeader.biWidth * bmiTile.bmiHeader.biHeight - 1)
    
    If 0 = GetDIBits(hDC, Tile.handle, 0, Abs(bmiTile.bmiHeader.biHeight), DIBits(0), bmiTile, DIB_RGB_COLORS) Then
        ReleaseDC 0, hDC
        Exit Sub
    End If
    
    ReDim DIBitsPic(0 To bmiPic.bmiHeader.biWidth * bmiTile.bmiHeader.biHeight - 1)
    If bmiPic.bmiHeader.biHeight > 0 Then
        If 0 = GetDIBits(hDC, Pic.handle, bmiPic.bmiHeader.biHeight - Y - Abs(bmiTile.bmiHeader.biHeight), Abs(bmiTile.bmiHeader.biHeight), DIBitsPic(0), bmiPic, DIB_RGB_COLORS) Then
            ReleaseDC 0, hDC
            Exit Sub
        End If
    Else
        If 0 = GetDIBits(hDC, Pic.handle, Y, Abs(bmiTile.bmiHeader.biHeight), DIBitsPic(0), bmiPic, DIB_RGB_COLORS) Then
            ReleaseDC 0, hDC
            Exit Sub
        End If
    End If
    
    For YIndex = 0 To Abs(bmiTile.bmiHeader.biHeight) - 1
        OffTile = YIndex * bmiTile.bmiHeader.biWidth
        If Sgn(bmiPic.bmiHeader.biHeight) <> Sgn(bmiTile.bmiHeader.biHeight) Then
            OffPic = (Abs(bmiTile.bmiHeader.biHeight) - 1 - YIndex) * bmiPic.bmiHeader.biWidth
        Else
            OffPic = YIndex * bmiPic.bmiHeader.biWidth
        End If
        For XIndex = 0 To bmiTile.bmiHeader.biWidth - 1
            DIBitsPic(OffPic + XIndex + X) = DIBits(OffTile + XIndex)
        Next
    Next
    
    If bmiPic.bmiHeader.biHeight > 0 Then
        SetDIBits hDC, Pic.handle, bmiPic.bmiHeader.biHeight - Abs(bmiTile.bmiHeader.biHeight) - Y, Abs(bmiTile.bmiHeader.biHeight), DIBitsPic(0), bmiPic, DIB_RGB_COLORS
    Else
        SetDIBits hDC, Pic.handle, Y, Abs(bmiTile.bmiHeader.biHeight), DIBitsPic(0), bmiPic, DIB_RGB_COLORS
    End If
    
    ReleaseDC 0, hDC
    
End Sub

Public Function ExtractTile(ByVal Pic As StdPicture, ByVal ExLeft As Long, ByVal ExTop As Long, ByVal ExWidth As Long, ByVal ExHeight As Long, Optional ByVal bHighlight As Boolean = False) As StdPicture
Attribute ExtractTile.VB_Description = "Extract a picture object from a larger picture object, optionally highlighting it in blue."
    Dim bmi As BITMAPINFO
    Dim hDC As Long
    Dim pBits As Long
    Dim DIBits() As RGBQUAD
    Dim hDib As Long
    Dim Y As Long
    Dim Wid As Long
    Dim X As Long
    Dim picGuid As IID
    Dim picDesc As PICTDESC
    Dim PicResult As IPictureDisp
    
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = ExWidth
        .biHeight = -ExHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    hDC = GetDC(0)
    If hDC = 0 Then
        Err.Raise vbObjectError, , "Cannot get device context to extract tile"
        Exit Function
    End If
    
    hDib = CreateDIBSection(hDC, bmi, DIB_RGB_COLORS, pBits, 0, 0)
    If hDib = 0 Then
        ReleaseDC 0, hDC
        Err.Raise Err.LastDllError
    End If
    
    bmi.bmiHeader.biBitCount = 0
    If 0 = GetDIBits(hDC, Pic.handle, ExTop, ExHeight, ByVal 0&, bmi, DIB_RGB_COLORS) Then
        DeleteObject hDib
        ReleaseDC 0, hDC
        Exit Function
    End If
    
    With bmi.bmiHeader
        Wid = .biWidth
        ReDim DIBits(0 To Wid * ExHeight - 1)
        .biBitCount = 32
        .biCompression = BI_RGB
    
        If .biHeight > 0 Then
            If ExHeight > GetDIBits(hDC, Pic.handle, .biHeight - ExTop - ExHeight, ExHeight, DIBits(0), bmi, DIB_RGB_COLORS) Then
                DeleteObject hDib
                ReleaseDC 0, hDC
                Exit Function
            End If
        Else
            If ExHeight > GetDIBits(hDC, Pic.handle, ExTop, ExHeight, DIBits(0), bmi, DIB_RGB_COLORS) Then
                DeleteObject hDib
                ReleaseDC 0, hDC
                Exit Function
            End If
        End If
    End With
    
    With bmi.bmiHeader
        .biBitCount = 32
        .biWidth = ExWidth
        .biHeight = -ExHeight
        .biCompression = BI_RGB
    End With
    
    For Y = 0 To ExHeight - 1
        If bHighlight Then
            For X = ExLeft To ExLeft + ExWidth - 1
                With DIBits(Wid * Y + X)
                    .rgbBlue = (.rgbBlue * 2 + 255) / 3
                    .rgbGreen = .rgbGreen * 2 / 3
                    .rgbRed = .rgbRed * 2 / 3
                End With
            Next
        End If
        If 0 = SetDIBits(hDC, hDib, Y, 1, DIBits(Wid * Y + ExLeft), bmi, DIB_RGB_COLORS) Then
            DeleteObject hDib
            ReleaseDC 0, hDC
            Exit Function
        End If
    Next
    
    picGuid = GetPicGuid
    
    ReleaseDC 0, hDC
    
    picDesc.cbSizeOfStruct = Len(picDesc)
    picDesc.hBitmap = hDib
    picDesc.hpal = 0
    picDesc.picType = PICTYPE_BITMAP
    
    If OleCreatePictureIndirect(picDesc, picGuid, True, PicResult) <> S_OK Then
        DeleteObject hDib
        Exit Function
    End If
    
    Set ExtractTile = PicResult
    Set PicResult = Nothing
    
End Function

Private Function GetPicGuid() As IID
    ' IID_IPictureDisp
    With GetPicGuid
        .X = &H7BF80981
        .s1 = &HBF32
        .s2 = &H101A
        .C(0) = &H8B
        .C(1) = &HBB
        .C(2) = &H0
        .C(3) = &HAA
        .C(4) = &H0
        .C(5) = &H30
        .C(6) = &HC
        .C(7) = &HAB
    End With
End Function

Public Function GetEntirePath(Path As String) As String
Attribute GetEntirePath.VB_Description = "Return a fully qualified path and file name based on a file name."
    Dim Buf As String
    Dim L As Long
    Dim I As Long
    
    Buf = Space$(256)
    
    L = GetFullPathName(Path & Chr$(0), Len(Buf), Buf, I)
    Buf = Left$(Buf, L)
    GetEntirePath = Buf
    
End Function

Public Function GetMouseBmp() As StdPicture
Attribute GetMouseBmp.VB_Description = "Global way to return a picture for mouse pointers on a full screen display."
    'Dim S As String
    
    Set GetMouseBmp = frmProject.imgMouse.Picture
    
    'S = App.Path
    'If Right$(S, 1) <> "\" Then S = S & "\"
    'S = S & "mouse.bmp"
    'GetMouseBmp = S
    
End Function

' Convert Upside-down rectangular coordinates to right-side-up polar coordinates
Public Sub RectToPolar(ByVal X As Single, ByVal Y As Single, ByRef Angle As Single, ByRef Distance As Single)
Attribute RectToPolar.VB_Description = "Convert rectangular coordinates to polar coordinates."
   If X <> 0 Then
       Angle = Atn(-Y / X)
   Else
       Angle = -(Pi / 2) * Sgn(Y)
   End If
   If X < 0 Then
       Angle = Pi + Angle
   ElseIf Y > 0 Then
       Angle = Pi * 2 + Angle
   End If
   Distance = Sqr(X * X + Y * Y)
End Sub

Public Sub SaveString(ByVal nFileNum As Integer, ByVal strVal As String)
Attribute SaveString.VB_Description = "Save a string into a binary file, first writing a 4-byte length."
    Dim SLen As Long
   
   SLen = Len(strVal)
   Put #nFileNum, , SLen
   If SLen > 0 Then
       Put #nFileNum, , strVal
   End If
   
End Sub

Public Function LoadString(ByVal nFileNum As Integer) As String
Attribute LoadString.VB_Description = "Load (and return) a string value from a binary file, first reading a 4-byte length, then the string value."
   Dim SLen As Long
   Dim S As String

   Get #nFileNum, , SLen
   If SLen > 0 Then
       S = Space$(SLen)
       Get #nFileNum, , S
       LoadString = S
   End If
   
End Function

Public Function DotProduct(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
Attribute DotProduct.VB_Description = "Compute the dot product of two vectors."
    If Abs(X1) + Abs(Y1) = 0 Or Abs(X2) + Abs(Y2) = 0 Then
        DotProduct = 0
        Exit Function
    End If
    DotProduct = (X1 * X2 + Y1 * Y2) / (Sqr(X1 * X1 + Y1 * Y1) * Sqr(X2 * X2 + Y2 * Y2))
End Function

Public Function SaveScreenShot() As String
    Dim ScrIdx As Integer
    
    For ScrIdx = 0 To 99
        If Len(Dir$(App.Path & "\Screen" & Format$(ScrIdx, "00") & ".bmp")) = 0 Then
            SavePicture CurDisp.ScreenShot, App.Path & "\Screen" & Format$(ScrIdx, "00") & ".bmp"
            SaveScreenShot = "Screen saved as " & App.Path & "\Screen" & Format$(ScrIdx, "00") & ".bmp"
            Exit For
        End If
    Next
End Function

Public Function Prj2XML() As String
Attribute Prj2XML.VB_Description = "Save project and map data into an XML file."
    Dim I As Integer, J As Integer, K As Integer, L As Integer
    Dim oLevel1Node As Object
    Dim oLevel2Node As Object
    Dim oLevel3Node As Object
    Dim oLevel4Node As Object
    Dim oLevel5Node As Object
    Dim oLevel6Node As Object
    Dim oLevel7Node As Object
    Dim oLevel8Node As Object
    Dim oRoot As Object
    Dim XMLDoc As Object

    On Error Resume Next
    Set XMLDoc = CreateObject("MSXML.DOMDocument")
    If Err.Number <> 0 Then
        MsgBox "Unable to create XML document.  MS XML support may not be installed.", vbExclamation
        Exit Function
    End If
    On Error GoTo XMLErr
    XMLDoc.resolveExternals = True
    Set oLevel1Node = XMLDoc.createProcessingInstruction("xml", "version='1.0'")
    Set oLevel1Node = XMLDoc.insertBefore(oLevel1Node, XMLDoc.childNodes.Item(0))
    Set oRoot = XMLDoc.createElement("GameProject")
    Set XMLDoc.documentElement = oRoot
    oRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
    oRoot.setAttribute "Version", "1"
    Set oLevel1Node = XMLDoc.createElement("Tilesets")
    oLevel1Node.setAttribute "Count", Prj.TileSetDefCount
    For I = 0 To Prj.TileSetDefCount - 1
       Set oLevel2Node = XMLDoc.createElement("Tileset")
       oLevel2Node.setAttribute "Index", CStr(I)
       With Prj.TileSetDef(I)
          AppendText XMLDoc, oLevel2Node, "Name", .Name
          AppendText XMLDoc, oLevel2Node, "ImagePath", .ImagePath
          AppendInt XMLDoc, oLevel2Node, "TileWidth", .TileWidth
          AppendInt XMLDoc, oLevel2Node, "TileHeight", .TileHeight
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    Set oLevel1Node = XMLDoc.createElement("Maps")
    oLevel1Node.setAttribute "Count", Prj.MapCount
    For I = 0 To Prj.MapCount - 1
       Set oLevel2Node = XMLDoc.createElement("Map")
       oLevel2Node.setAttribute "Index", CStr(I)
       With Prj.Maps(I)
          AppendText XMLDoc, oLevel2Node, "Name", .Name
          AppendText XMLDoc, oLevel2Node, "Path", .Path
          AppendInt XMLDoc, oLevel2Node, "ViewLeft", .ViewLeft
          AppendInt XMLDoc, oLevel2Node, "ViewTop", .ViewTop
          AppendInt XMLDoc, oLevel2Node, "ViewWidth", .ViewWidth
          AppendInt XMLDoc, oLevel2Node, "ViewHeight", .ViewHeight
          AppendInt XMLDoc, oLevel2Node, "MapWidth", .MapWidth
          AppendInt XMLDoc, oLevel2Node, "MapHeight", .MapHeight
          AppendText XMLDoc, oLevel2Node, "BackgroundMusic", .BackgroundMusic
          AppendText XMLDoc, oLevel2Node, "PlayerSpriteName", .PlayerSpriteName
          Set oLevel3Node = XMLDoc.createElement("Layers")
          oLevel3Node.setAttribute "Count", .LayerCount
          For J = 0 To .LayerCount - 1
             With .MapLayer(J)
                Set oLevel4Node = XMLDoc.createElement("Layer")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendText XMLDoc, oLevel4Node, "Name", .Name
                AppendText XMLDoc, oLevel4Node, "TSDefName", .TSDef.Name
                AppendFloat XMLDoc, oLevel4Node, "XScrollRate", .XScrollRate
                AppendFloat XMLDoc, oLevel4Node, "YScrollRate", .YScrollRate
                AppendBoolean XMLDoc, oLevel4Node, "Transparent", .Transparent
                AppendInt XMLDoc, oLevel4Node, "Columns", .Columns
                AppendInt XMLDoc, oLevel4Node, "Rows", .Rows
                AppendBinary XMLDoc, oLevel4Node, "Tiles", .Data.MapData
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("Paths")
          oLevel3Node.setAttribute "Count", .PathCount
          For J = 0 To .PathCount - 1
             With .Paths(J)
                Set oLevel4Node = XMLDoc.createElement("Path")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendText XMLDoc, oLevel4Node, "Name", .Name
                AppendText XMLDoc, oLevel4Node, "LayerName", .LayerName
                Set oLevel5Node = XMLDoc.createElement("Points")
                oLevel5Node.setAttribute "Count", .PointCount
                For K = 0 To .PointCount - 1
                   Set oLevel6Node = XMLDoc.createElement("Point")
                   oLevel6Node.setAttribute "Index", CStr(K)
                   AppendInt XMLDoc, oLevel6Node, "X", .PointX(K)
                   AppendInt XMLDoc, oLevel6Node, "Y", .PointY(K)
                   oLevel5Node.appendChild oLevel6Node
                Next
                oLevel4Node.appendChild oLevel5Node
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("Templates")
          oLevel3Node.setAttribute "Count", .SpriteTemplateCount
          For J = 0 To .SpriteTemplateCount - 1
             With .SpriteTemplates(J)
                Set oLevel4Node = XMLDoc.createElement("Template")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendText XMLDoc, oLevel4Node, "Name", .Name
                If Not (.SolidInfo Is Nothing) Then
                   AppendText XMLDoc, oLevel4Node, "SolidName", .SolidInfo.Name
                   AppendText XMLDoc, oLevel4Node, "SolidTS", .SolidInfo.TSName
                End If
                AppendInt XMLDoc, oLevel4Node, "Flags", .Flags
                AppendInt XMLDoc, oLevel4Node, "AnimSpeed", .AnimSpeed
                AppendInt XMLDoc, oLevel4Node, "MoveSpeed", .MoveSpeed
                AppendInt XMLDoc, oLevel4Node, "GravPow", .GravPow
                AppendInt XMLDoc, oLevel4Node, "Inertia", .Inertia
                AppendInt XMLDoc, oLevel4Node, "CollClass", .CollClass
                AppendInt XMLDoc, oLevel4Node, "StateType", .StateType
                AppendInt XMLDoc, oLevel4Node, "ControlType", .ControlType
                Set oLevel5Node = XMLDoc.createElement("States")
                oLevel5Node.setAttribute "Count", .StateCount
                For K = 0 To .StateCount - 1
                   Set oLevel6Node = XMLDoc.createElement("State")
                   oLevel6Node.setAttribute "Index", CStr(K)
                   If Not (.StateTilesetDef(K) Is Nothing) Then
                      AppendInt XMLDoc, oLevel6Node, "TSIdx", Prj.TileSetDef(.StateTilesetDef(K).Name).GetIndex
                   End If
                   Set oLevel7Node = XMLDoc.createElement("Frames")
                   oLevel7Node.setAttribute "Count", .StateFrameCount(K)
                   For L = 0 To .StateFrameCount(K) - 1
                      Set oLevel8Node = XMLDoc.createElement("Frame")
                      oLevel8Node.setAttribute "Index", CStr(L)
                      oLevel8Node.dataType = "ui1"
                      oLevel8Node.nodeTypedValue = .StateFrame(K, L)
                      oLevel7Node.appendChild oLevel8Node
                   Next
                   oLevel6Node.appendChild oLevel7Node
                   oLevel5Node.appendChild oLevel6Node
                Next
                oLevel4Node.appendChild oLevel5Node
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("SpriteDefs")
          oLevel3Node.setAttribute "Count", .SpriteDefCount
          For J = 0 To .SpriteDefCount - 1
             With .SpriteDefs(J)
                Set oLevel4Node = XMLDoc.createElement("SpriteDef")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendText XMLDoc, oLevel4Node, "Name", .Name
                AppendText XMLDoc, oLevel4Node, "LayerName", .rLayer.Name
                AppendInt XMLDoc, oLevel4Node, "Flags", .Flags
                AppendText XMLDoc, oLevel4Node, "Path", .rPath.Name
                If Not (.Template Is Nothing) Then
                   AppendInt XMLDoc, oLevel4Node, "TemplateIndex", .Template.GetIndexes(CInt(K))
                End If
                oLevel3Node.appendChild oLevel4Node
             End With
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("CollDefs")
          oLevel3Node.setAttribute "Count", .CollDefCount
          For J = 0 To .CollDefCount - 1
             With .CollDefs(J)
                Set oLevel4Node = XMLDoc.createElement("CollDef")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendInt XMLDoc, oLevel4Node, "ClassA", .ClassA
                AppendInt XMLDoc, oLevel4Node, "ClassB", .ClassB
                AppendInt XMLDoc, oLevel4Node, "Flags", .Flags
                AppendText XMLDoc, oLevel4Node, "SpecialFunction", .SpecialFunction
                AppendInt XMLDoc, oLevel4Node, "InvFlags", .InvFlags
                AppendInt XMLDoc, oLevel4Node, "InvItem", .InvItem
                AppendInt XMLDoc, oLevel4Node, "InvUseCount", .InvUseCount
                AppendText XMLDoc, oLevel4Node, "Media", .Media
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("CollClasses")
          For J = 0 To 15
             Set oLevel4Node = XMLDoc.createElement("CollClass")
             oLevel4Node.setAttribute "Index", CStr(J)
             AppendText XMLDoc, oLevel4Node, "Name", .CollClassName(J)
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("Specials")
          oLevel3Node.setAttribute "Count", .SpecialCount
          For J = 0 To .SpecialCount - 1
             With .Specials(J)
                Set oLevel4Node = XMLDoc.createElement("Special")
                oLevel4Node.setAttribute "Index", CStr(J)
                AppendText XMLDoc, oLevel4Node, "Name", .Name
                AppendInt XMLDoc, oLevel4Node, "LayerIndex", .LayerIndex
                AppendInt XMLDoc, oLevel4Node, "TileLeft", .TileLeft
                AppendInt XMLDoc, oLevel4Node, "TileTop", .TileTop
                AppendInt XMLDoc, oLevel4Node, "TileRight", .TileRight
                AppendInt XMLDoc, oLevel4Node, "TileBottom", .TileBottom
                AppendInt XMLDoc, oLevel4Node, "Flags", .Flags
                AppendInt XMLDoc, oLevel4Node, "FuncType", .FuncType
                AppendText XMLDoc, oLevel4Node, "Value", .Value
                AppendText XMLDoc, oLevel4Node, "SpriteName", .SpriteName
                AppendInt XMLDoc, oLevel4Node, "DestX", .DestX
                AppendInt XMLDoc, oLevel4Node, "DestY", .DestY
                AppendText XMLDoc, oLevel4Node, "MediaName", .MediaName
                AppendInt XMLDoc, oLevel4Node, "InvItem", .InvItem
                AppendInt XMLDoc, oLevel4Node, "InvUseCount", .InvUseCount
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
          Set oLevel3Node = XMLDoc.createElement("Interactions")
          oLevel3Node.setAttribute "Count", .InteractCount
          For J = 0 To .InteractCount - 1
             With .Interactions(J)
                Set oLevel4Node = XMLDoc.createElement("Interaction")
                oLevel4Node.setAttribute "Index", CStr(J)
                If Not (.TouchCategory Is Nothing) Then
                   AppendText XMLDoc, oLevel4Node, "CategoryName", .TouchCategory.Name
                   AppendText XMLDoc, oLevel4Node, "TSName", .TouchCategory.TSName
                End If
                AppendInt XMLDoc, oLevel4Node, "Flags", .Flags
                AppendInt XMLDoc, oLevel4Node, "Reaction", .Reaction
                AppendInt XMLDoc, oLevel4Node, "InvItem", .InvItem
                AppendInt XMLDoc, oLevel4Node, "ReplaceTile", .ReplaceTile
                AppendText XMLDoc, oLevel4Node, "Media", .Media
             End With
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("MatchDefs")
    oLevel1Node.setAttribute "Count", Prj.MatchDefCount
    For I = 0 To Prj.MatchDefCount - 1
       With Prj.MatchDefs(I)
          Set oLevel2Node = XMLDoc.createElement("MatchDef")
          oLevel2Node.setAttribute "Index", CStr(I)
          AppendText XMLDoc, oLevel2Node, "Name", .Name
          AppendText XMLDoc, oLevel2Node, "MatchGroup", .AllTiles.Serialize
          AppendText XMLDoc, oLevel2Node, "TileSet", .TSDef.Name
          AppendText XMLDoc, oLevel2Node, "TLGroup", .TileMatches.MatchGroup(0).Serialize
          AppendText XMLDoc, oLevel2Node, "TGroup", .TileMatches.MatchGroup(1).Serialize
          AppendText XMLDoc, oLevel2Node, "TRGroup", .TileMatches.MatchGroup(2).Serialize
          AppendText XMLDoc, oLevel2Node, "ITLGroup", .TileMatches.MatchGroup(3).Serialize
          AppendText XMLDoc, oLevel2Node, "ITRGroup", .TileMatches.MatchGroup(4).Serialize
          AppendText XMLDoc, oLevel2Node, "LGroup", .TileMatches.MatchGroup(5).Serialize
          AppendText XMLDoc, oLevel2Node, "CGroup", .TileMatches.MatchGroup(6).Serialize
          AppendText XMLDoc, oLevel2Node, "RGroup", .TileMatches.MatchGroup(7).Serialize
          AppendText XMLDoc, oLevel2Node, "IBLGroup", .TileMatches.MatchGroup(8).Serialize
          AppendText XMLDoc, oLevel2Node, "IBRGroup", .TileMatches.MatchGroup(9).Serialize
          AppendText XMLDoc, oLevel2Node, "BLGroup", .TileMatches.MatchGroup(10).Serialize
          AppendText XMLDoc, oLevel2Node, "BGroup", .TileMatches.MatchGroup(11).Serialize
          AppendText XMLDoc, oLevel2Node, "BRGroup", .TileMatches.MatchGroup(12).Serialize
          AppendText XMLDoc, oLevel2Node, "IDRGroup", .TileMatches.MatchGroup(13).Serialize
          AppendText XMLDoc, oLevel2Node, "IDLGroup", .TileMatches.MatchGroup(14).Serialize
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("AnimDefs")
    oLevel1Node.setAttribute "Count", Prj.AnimDefCount
    For I = 0 To Prj.AnimDefCount - 1
       With Prj.AnimDefs(I)
          Set oLevel2Node = XMLDoc.createElement("AnimDef")
          oLevel2Node.setAttribute "Index", CStr(I)
          AppendText XMLDoc, oLevel2Node, "Name", .Name
          AppendText XMLDoc, oLevel2Node, "MapName", .MapName
          AppendText XMLDoc, oLevel2Node, "LayerName", .LayerName
          AppendInt XMLDoc, oLevel2Node, "BaseTile", .BaseTile
          Set oLevel3Node = XMLDoc.createElement("Frames")
          oLevel3Node.setAttribute "Count", .FrameCount
          For J = 0 To .FrameCount - 1
             Set oLevel4Node = XMLDoc.createElement("Frame")
             oLevel4Node.setAttribute "Index", CStr(J)
             AppendInt XMLDoc, oLevel4Node, "Tile", .FrameValue(J)
             AppendInt XMLDoc, oLevel4Node, "Delay", .FrameDelay(J)
             oLevel3Node.appendChild oLevel4Node
          Next
          oLevel2Node.appendChild oLevel3Node
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("Categories")
    oLevel1Node.setAttribute "Count", Prj.GroupCount
    For I = 0 To Prj.GroupCount - 1
       With Prj.GroupByIndex(I)
          Set oLevel2Node = XMLDoc.createElement("Category")
          oLevel2Node.setAttribute "Index", CStr(I)
          AppendText XMLDoc, oLevel2Node, "GroupName", .Name
          AppendText XMLDoc, oLevel2Node, "TilesetName", .TSName
          AppendText XMLDoc, oLevel2Node, "Group", .Group.Serialize
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("SolidDefs")
    oLevel1Node.setAttribute "Count", Prj.SolidDefCount
    For I = 0 To Prj.SolidDefCount - 1
       With Prj.SolidDefsByIndex(I)
          Set oLevel2Node = XMLDoc.createElement("SolidDef")
          oLevel2Node.setAttribute "Index", CStr(I)
          AppendText XMLDoc, oLevel2Node, "SolidDefName", .Name
          AppendText XMLDoc, oLevel2Node, "TileSetName", .TSName
          If Not .Solid Is Nothing Then
             AppendText XMLDoc, oLevel2Node, "Solid", .Solid.Name
          End If
          If Not .Uphill Is Nothing Then
             AppendText XMLDoc, oLevel2Node, "Uphill", .Uphill.Name
          End If
          If Not .Downhill Is Nothing Then
             AppendText XMLDoc, oLevel2Node, "Downhill", .Downhill.Name
          End If
          If Not .UpCeil Is Nothing Then
             AppendText XMLDoc, oLevel2Node, "UpCeil", .UpCeil.Name
          End If
          If Not .DownCeil Is Nothing Then
             AppendText XMLDoc, oLevel2Node, "DownCeil", .DownCeil.Name
          End If
       End With
       oLevel1Node.appendChild oLevel2Node
    Next
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("Player")
    With Prj.GamePlayer
       AppendInt XMLDoc, oLevel1Node, "ScrollMarginX", .ScrollMarginX
       AppendInt XMLDoc, oLevel1Node, "ScrollMarginY", .ScrollMarginY
       AppendInt XMLDoc, oLevel1Node, "InvMargin", .InvBarMargin
       AppendText XMLDoc, oLevel1Node, "StartMap", .StartMapName
       AppendInt XMLDoc, oLevel1Node, "KeyUp", .KeyConfig(0)
       AppendInt XMLDoc, oLevel1Node, "KeyLeft", .KeyConfig(1)
       AppendInt XMLDoc, oLevel1Node, "KeyRight", .KeyConfig(2)
       AppendInt XMLDoc, oLevel1Node, "KeyDown", .KeyConfig(3)
       AppendInt XMLDoc, oLevel1Node, "KeyBtn1", .KeyConfig(4)
       AppendInt XMLDoc, oLevel1Node, "KeyBtn2", .KeyConfig(5)
       AppendInt XMLDoc, oLevel1Node, "KeyBtn3", .KeyConfig(6)
       AppendInt XMLDoc, oLevel1Node, "KeyBtn4", .KeyConfig(7)
       AppendBoolean XMLDoc, oLevel1Node, "EnableJoystick", .bEnableJoystick
       Set oLevel2Node = XMLDoc.createElement("Inventory")
       oLevel2Node.setAttribute "Count", .InventoryCount
       For I = 0 To .InventoryCount - 1
          Set oLevel3Node = XMLDoc.createElement("Item")
          oLevel3Node.setAttribute "Index", CStr(I)
          AppendText XMLDoc, oLevel3Node, "Name", .InventoryItemName(I)
          AppendText XMLDoc, oLevel3Node, "IconTilesetName", Prj.TileSetDef(.InvIconTilesetIdx(I)).Name
          AppendInt XMLDoc, oLevel3Node, "IconTileIndex", .InvIconTileIdx(I)
          AppendInt XMLDoc, oLevel3Node, "QuantityOwned", .InvQuantityOwned(I)
          AppendInt XMLDoc, oLevel3Node, "MaxQuantity", .InvMaxQuantity(I)
          AppendInt XMLDoc, oLevel3Node, "QuantityDisplayType", .InvQuantityDisplayType(I)
          AppendInt XMLDoc, oLevel3Node, "IconCountPerRepeat", .InvIconCountPerRepeat(I)
          AppendInt XMLDoc, oLevel3Node, "BarColor", .InvBarColor(I)
          AppendInt XMLDoc, oLevel3Node, "BarBackgroundColor", .InvBarBackgroundColor(I)
          AppendInt XMLDoc, oLevel3Node, "BarOutlineColor", .InvBarOutlineColor(I)
          AppendInt XMLDoc, oLevel3Node, "BarLength", .InvBarLength(I)
          AppendInt XMLDoc, oLevel3Node, "BarThickness", .InvBarThickness(I)
          AppendInt XMLDoc, oLevel3Node, "DisplayX", .InvDisplayX(I)
          AppendInt XMLDoc, oLevel3Node, "DisplayY", .InvDisplayY(I)
          oLevel2Node.appendChild oLevel3Node
       Next
       oLevel1Node.appendChild oLevel2Node
    End With
    oRoot.appendChild oLevel1Node
    
    Set oLevel1Node = XMLDoc.createElement("Media")
    With Prj.MediaMgr
       oLevel1Node.setAttribute "Count", .MediaClipCount
       For I = 0 To .MediaClipCount - 1
          With .Clip(I)
             Set oLevel2Node = XMLDoc.createElement("Clip")
             oLevel2Node.setAttribute "Index", CStr(I)
             AppendText XMLDoc, oLevel2Node, "Name", .Name
             AppendText XMLDoc, oLevel2Node, "MediaFile", .strMediaFile
             AppendInt XMLDoc, oLevel2Node, "OutputX", .OutputX
             AppendInt XMLDoc, oLevel2Node, "OutputY", .OutputY
             AppendInt XMLDoc, oLevel2Node, "Flags", .Flags
             AppendInt XMLDoc, oLevel2Node, "Volume", .Volume
          End With
          oLevel1Node.appendChild oLevel2Node
       Next
    End With
    oRoot.appendChild oLevel1Node
    
    Prj2XML = XMLDoc.xml
    
    Set oLevel8Node = Nothing
    Set oLevel7Node = Nothing
    Set oLevel6Node = Nothing
    Set oLevel5Node = Nothing
    Set oLevel4Node = Nothing
    Set oLevel3Node = Nothing
    Set oLevel2Node = Nothing
    Set oLevel1Node = Nothing
    Set oRoot = Nothing
    Set XMLDoc = Nothing
    Exit Function

XMLErr:
    MsgBox "Error exporting XML: " & Err.Description, vbExclamation
End Function

Private Sub AppendInt(ByRef oXMLDoc As Object, ByRef oParent As Object, ByVal strName As String, ByVal intVal As Long)
   Dim oNode

   Set oNode = oXMLDoc.createElement(strName)
   oNode.dataType = "int"
   oNode.nodeTypedValue = intVal
   oParent.appendChild oNode
End Sub

Private Sub AppendFloat(ByRef oXMLDoc As Object, ByRef oParent As Object, ByVal strName As String, ByVal dblVal As Double)
   Dim oNode

   Set oNode = oXMLDoc.createElement(strName)
   oNode.dataType = "float"
   oNode.nodeTypedValue = dblVal
   oParent.appendChild oNode
End Sub

Private Sub AppendText(ByRef oXMLDoc As Object, ByRef oParent As Object, ByVal strName As String, ByVal strVal As String)
   Dim oNode

   Set oNode = oXMLDoc.createElement(strName)
   oNode.appendChild oXMLDoc.createCDATASection(strVal)
   'oNode.Text = strVal
   oParent.appendChild oNode
End Sub

Private Sub AppendBinary(ByRef oXMLDoc As Object, ByRef oParent As Object, ByVal strName As String, arbyVal() As Byte)
   Dim oNode

   Set oNode = oXMLDoc.createElement(strName)
   oNode.dataType = "bin.base64"
   oNode.nodeTypedValue = arbyVal
   oParent.appendChild oNode
End Sub

Private Sub AppendBoolean(ByRef oXMLDoc As Object, ByRef oParent As Object, ByVal strName As String, ByVal bVal As Boolean)
   Dim oNode

   Set oNode = oXMLDoc.createElement(strName)
   oNode.dataType = "boolean"
   oNode.Text = IIf(bVal, "1", "0")
   oParent.appendChild oNode
End Sub

Public Function XML2Prj(strXML As String) As Boolean
Attribute XML2Prj.VB_Description = "Load a project and maps from an XML file"
    Dim I As Integer, J As Integer, K As Integer, L As Integer
    Dim oLevel1Node As Object
    Dim oLevel2Node As Object
    Dim oLevel3Node As Object
    Dim oLevel4Node As Object
    Dim oLevel5Node As Object
    Dim oLevel6Node As Object
    Dim oLevel7Node As Object
    Dim oLevel8Node As Object
    Dim oRoot As Object
    Dim XMLDoc As Object
    Dim NewMap As Map
    Dim NewPath As Path
    Dim NewTpl As SpriteTemplate
    Dim NewMatch As MatchDef
    Dim NewAnim As AnimDef
    Dim NewClip As MediaClip
    Dim NewSprDef As SpriteDef
    Dim NewCollDef As CollisionDef
    Dim NewSpecial As SpecialFunction
    Dim NewInteraction As Interaction
    Dim strKey As String
    Dim strKey2 As String

    On Error Resume Next
    Set XMLDoc = CreateObject("MSXML.DOMDocument")
    If Err.Number <> 0 Then
        MsgBox "Unable to create XML document.  MS XML support may not be installed.", vbExclamation
        XML2Prj = False
        Exit Function
    End If

    On Error GoTo LoadXMLErr
    XMLDoc.resolveExternals = True
    
    If Not (XMLDoc.loadXML(strXML)) Then
        MsgBox "Error loading XML: " & XMLDoc.parseError.reason, vbExclamation
        XML2Prj = False
        Exit Function
    End If

    Set oRoot = XMLDoc.selectSingleNode("/GameProject")
    If oRoot Is Nothing Then
        MsgBox "The XML file does not appear to contain a GameDev project.", vbExclamation
        XML2Prj = False
        Exit Function
    End If
    If oRoot.selectSingleNode("/GameProject/@Version").Value <> 1 Then
        MsgBox "The version of the GameDev project in the XML file must be version 2 for this version of GameDev.", vbExclamation
        XML2Prj = False
        Exit Function
    End If
    Set Prj = New GameProject
    Set oLevel1Node = oRoot.selectSingleNode("Tilesets")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("Tileset[@Index=" & CStr(I) & "]")
        Prj.AddTileSet oLevel2Node.selectSingleNode("ImagePath").nodeTypedValue, _
                       oLevel2Node.selectSingleNode("TileWidth").nodeTypedValue, _
                       oLevel2Node.selectSingleNode("TileHeight").nodeTypedValue, _
                       oLevel2Node.selectSingleNode("Name").nodeTypedValue
    Next
    
    Set oLevel1Node = oRoot.selectSingleNode("MatchDefs")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("MatchDef[@Index=" & CStr(I) & "]")
        Set NewMatch = New MatchDef
        NewMatch.Name = oLevel2Node.selectSingleNode("Name").nodeTypedValue
        NewMatch.AllTiles.Deserialize oLevel2Node.selectSingleNode("MatchGroup").nodeTypedValue
        strKey = oLevel2Node.selectSingleNode("TileSet").nodeTypedValue
        Set NewMatch.TSDef = Prj.TileSetDef(strKey)
        NewMatch.TileMatches.MatchGroup(0).Deserialize oLevel2Node.selectSingleNode("TLGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(1).Deserialize oLevel2Node.selectSingleNode("TGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(2).Deserialize oLevel2Node.selectSingleNode("TRGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(3).Deserialize oLevel2Node.selectSingleNode("ITLGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(4).Deserialize oLevel2Node.selectSingleNode("ITRGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(5).Deserialize oLevel2Node.selectSingleNode("LGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(6).Deserialize oLevel2Node.selectSingleNode("CGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(7).Deserialize oLevel2Node.selectSingleNode("RGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(8).Deserialize oLevel2Node.selectSingleNode("IBLGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(9).Deserialize oLevel2Node.selectSingleNode("IBRGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(10).Deserialize oLevel2Node.selectSingleNode("BLGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(11).Deserialize oLevel2Node.selectSingleNode("BGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(12).Deserialize oLevel2Node.selectSingleNode("BRGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(13).Deserialize oLevel2Node.selectSingleNode("IDRGroup").nodeTypedValue
        NewMatch.TileMatches.MatchGroup(14).Deserialize oLevel2Node.selectSingleNode("IDLGroup").nodeTypedValue
        Prj.AddMatch NewMatch
    Next
    Set NewMatch = Nothing
    
    Set oLevel1Node = oRoot.selectSingleNode("AnimDefs")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("AnimDef[@Index=" & CStr(I) & "]")
        Set NewAnim = New AnimDef
        NewAnim.Name = oLevel2Node.selectSingleNode("Name").nodeTypedValue
        NewAnim.MapName = oLevel2Node.selectSingleNode("MapName").nodeTypedValue
        NewAnim.LayerName = oLevel2Node.selectSingleNode("LayerName").nodeTypedValue
        NewAnim.BaseTile = oLevel2Node.selectSingleNode("BaseTile").nodeTypedValue
        Set oLevel3Node = oLevel2Node.selectSingleNode("Frames")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Frame[@Index=" & CStr(J) & "]")
            NewAnim.InsertFrame J, oLevel4Node.selectSingleNode("Tile").nodeTypedValue, _
                                   oLevel4Node.selectSingleNode("Delay").nodeTypedValue
        Next
        Prj.AddAnim NewAnim
    Next
    Set NewAnim = Nothing

    Set oLevel1Node = oRoot.selectSingleNode("Categories")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("Category[@Index=" & CStr(I) & "]")
        With Prj.AddGroup(oLevel2Node.selectSingleNode("GroupName").nodeTypedValue, _
                          oLevel2Node.selectSingleNode("TilesetName").nodeTypedValue)
            .Group.Deserialize oLevel2Node.selectSingleNode("Group").nodeTypedValue
        End With
    Next

    Set oLevel1Node = oRoot.selectSingleNode("SolidDefs")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("SolidDef[@Index=" & CStr(I) & "]")
        With Prj.AddSolidDef(oLevel2Node.selectSingleNode("SolidDefName").nodeTypedValue, _
                             oLevel2Node.selectSingleNode("TileSetName").nodeTypedValue)
            strKey = oLevel2Node.selectSingleNode("Solid").nodeTypedValue
            If Prj.GroupExists(strKey, .TSName) Then Set .Solid = Prj.Groups(strKey, .TSName)
            If oLevel2Node.selectNodes("Uphill").length > 0 Then
                strKey = oLevel2Node.selectSingleNode("Uphill").nodeTypedValue
                If Prj.GroupExists(strKey, .TSName) Then Set .Uphill = Prj.Groups(strKey, .TSName)
            End If
            If oLevel2Node.selectNodes("Downhill").length > 0 Then
                strKey = oLevel2Node.selectSingleNode("Downhill").nodeTypedValue
                If Prj.GroupExists(strKey, .TSName) Then Set .Downhill = Prj.Groups(strKey, .TSName)
            End If
            If oLevel2Node.selectNodes("UpCeil").length > 0 Then
                strKey = oLevel2Node.selectSingleNode("UpCeil").nodeTypedValue
                If Prj.GroupExists(strKey, .TSName) Then Set .UpCeil = Prj.Groups(strKey, .TSName)
            End If
            If oLevel2Node.selectNodes("DownCeil").length > 0 Then
                strKey = oLevel2Node.selectSingleNode("DownCeil").nodeTypedValue
                If Prj.GroupExists(strKey, .TSName) Then Set .DownCeil = Prj.Groups(strKey, .TSName)
            End If
        End With
    Next
    
    Set oLevel1Node = oRoot.selectSingleNode("Player")
    With Prj.GamePlayer
        .ScrollMarginX = oLevel1Node.selectSingleNode("ScrollMarginX").nodeTypedValue
        .ScrollMarginY = oLevel1Node.selectSingleNode("ScrollMarginY").nodeTypedValue
        .InvBarMargin = oLevel1Node.selectSingleNode("InvMargin").nodeTypedValue
        .StartMapName = oLevel1Node.selectSingleNode("StartMap").nodeTypedValue
        .KeyConfig(0) = oLevel1Node.selectSingleNode("KeyUp").nodeTypedValue
        .KeyConfig(1) = oLevel1Node.selectSingleNode("KeyLeft").nodeTypedValue
        .KeyConfig(2) = oLevel1Node.selectSingleNode("KeyRight").nodeTypedValue
        .KeyConfig(3) = oLevel1Node.selectSingleNode("KeyDown").nodeTypedValue
        .KeyConfig(4) = oLevel1Node.selectSingleNode("KeyBtn1").nodeTypedValue
        .KeyConfig(5) = oLevel1Node.selectSingleNode("KeyBtn2").nodeTypedValue
        .KeyConfig(6) = oLevel1Node.selectSingleNode("KeyBtn3").nodeTypedValue
        .KeyConfig(7) = oLevel1Node.selectSingleNode("KeyBtn4").nodeTypedValue
        .bEnableJoystick = oLevel1Node.selectSingleNode("EnableJoystick").nodeTypedValue
        Set oLevel2Node = oLevel1Node.selectSingleNode("Inventory")
        For I = 0 To oLevel2Node.selectSingleNode("@Count").Value - 1
            Set oLevel3Node = oLevel2Node.selectSingleNode("Item[@Index=" & CStr(I) & "]")
            .AddInventoryItem oLevel3Node.selectSingleNode("QuantityDisplayType").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("DisplayX").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("DisplayY").nodeTypedValue
            .InventoryItemName(I) = oLevel3Node.selectSingleNode("Name").nodeTypedValue
            .SetInventoryTile I, oLevel3Node.selectSingleNode("IconTilesetName").nodeTypedValue, _
                                 oLevel3Node.selectSingleNode("IconTileIndex").nodeTypedValue
            .InvQuantityOwned(I) = oLevel3Node.selectSingleNode("QuantityOwned").nodeTypedValue
            .InvMaxQuantity(I) = oLevel3Node.selectSingleNode("MaxQuantity").nodeTypedValue
            .InvIconCountPerRepeat(I) = oLevel3Node.selectSingleNode("IconCountPerRepeat").nodeTypedValue
            .InvSetBarInfo I, oLevel3Node.selectSingleNode("BarColor").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("BarThickness").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("BarLength").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("BarBackgroundColor").nodeTypedValue, _
                              oLevel3Node.selectSingleNode("BarOutlineColor").nodeTypedValue
        Next
    End With

    Set oLevel1Node = oRoot.selectSingleNode("Media")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("Clip[@Index=" & CStr(I) & "]")
        Set NewClip = New MediaClip
        NewClip.Name = oLevel2Node.selectSingleNode("Name").nodeTypedValue
        NewClip.strMediaFile = oLevel2Node.selectSingleNode("MediaFile").nodeTypedValue
        NewClip.OutputX = oLevel2Node.selectSingleNode("OutputX").nodeTypedValue
        NewClip.OutputY = oLevel2Node.selectSingleNode("OutputY").nodeTypedValue
        NewClip.Flags = oLevel2Node.selectSingleNode("Flags").nodeTypedValue
        NewClip.Volume = oLevel2Node.selectSingleNode("Volume").nodeTypedValue
        Prj.MediaMgr.AddClip NewClip
    Next
    Set NewClip = Nothing

    Set oLevel1Node = oRoot.selectSingleNode("Maps")
    For I = 0 To oLevel1Node.selectSingleNode("@Count").Value - 1
        Set oLevel2Node = oLevel1Node.selectSingleNode("Map[@Index=" & CStr(I) & "]")
        Set NewMap = New Map
        NewMap.Name = oLevel2Node.selectSingleNode("Name").nodeTypedValue
        NewMap.Path = oLevel2Node.selectSingleNode("Path").nodeTypedValue
        NewMap.ViewLeft = oLevel2Node.selectSingleNode("ViewLeft").nodeTypedValue
        NewMap.ViewTop = oLevel2Node.selectSingleNode("ViewTop").nodeTypedValue
        NewMap.ViewWidth = oLevel2Node.selectSingleNode("ViewWidth").nodeTypedValue
        NewMap.ViewHeight = oLevel2Node.selectSingleNode("ViewHeight").nodeTypedValue
        NewMap.MapWidth = oLevel2Node.selectSingleNode("MapWidth").nodeTypedValue
        NewMap.MapHeight = oLevel2Node.selectSingleNode("MapHeight").nodeTypedValue
        NewMap.BackgroundMusic = oLevel2Node.selectSingleNode("BackgroundMusic").nodeTypedValue
        NewMap.PlayerSpriteName = oLevel2Node.selectSingleNode("PlayerSpriteName").nodeTypedValue
        Prj.AddMap NewMap
        Set oLevel3Node = oLevel2Node.selectSingleNode("Layers")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Layer[@Index=" & CStr(J) & "]")
            NewMap.AddLayer oLevel4Node.selectSingleNode("Name").nodeTypedValue, _
                            oLevel4Node.selectSingleNode("TSDefName").nodeTypedValue, _
                            oLevel4Node.selectSingleNode("XScrollRate").nodeTypedValue, _
                            oLevel4Node.selectSingleNode("YScrollRate").nodeTypedValue, _
                            oLevel4Node.selectSingleNode("Transparent").nodeTypedValue
            If NewMap.MapLayer(J).Columns <> oLevel4Node.selectSingleNode("Columns").nodeTypedValue Or _
               NewMap.MapLayer(J).Rows <> oLevel4Node.selectSingleNode("Rows").nodeTypedValue Then
                MsgBox "Layer data inconsistend with map size.  Aborting XML import.", vbExclamation
                Exit Function
            End If
            NewMap.MapLayer(J).Data.MapData = oLevel4Node.selectSingleNode("Tiles").nodeTypedValue
        Next
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("Paths")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Path[@Index=" & CStr(J) & "]")
            Set NewPath = New Path
            NewPath.Name = oLevel4Node.selectSingleNode("Name").nodeTypedValue
            NewPath.LayerName = oLevel4Node.selectSingleNode("LayerName").nodeTypedValue
            Set oLevel5Node = oLevel4Node.selectSingleNode("Points")
            For K = 0 To oLevel5Node.selectSingleNode("@Count").Value - 1
                Set oLevel6Node = oLevel5Node.selectSingleNode("Point[@Index=" & CStr(K) & "]")
                NewPath.AddPoint oLevel6Node.selectSingleNode("X").nodeTypedValue, _
                                 oLevel6Node.selectSingleNode("Y").nodeTypedValue
            Next
            NewMap.AddPath NewPath
        Next
        Set NewPath = Nothing
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("Templates")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Template[@Index=" & CStr(J) & "]")
            Set NewTpl = New SpriteTemplate
            NewTpl.Name = oLevel4Node.selectSingleNode("Name").nodeTypedValue
            If oLevel4Node.selectNodes("SolidName").length > 0 Then
                strKey = oLevel4Node.selectSingleNode("SolidName").nodeTypedValue
                strKey2 = oLevel4Node.selectSingleNode("SolidTS").nodeTypedValue
            Else
                strKey = ""
                strKey2 = ""
            End If
            If Len(strKey) > 0 And Len(strKey2) > 0 Then
                Set NewTpl.SolidInfo = Prj.SolidDefs(strKey, strKey2)
            Else
                Set NewTpl.SolidInfo = Nothing
            End If
            NewTpl.Flags = oLevel4Node.selectSingleNode("Flags").nodeTypedValue
            NewTpl.AnimSpeed = oLevel4Node.selectSingleNode("AnimSpeed").nodeTypedValue
            NewTpl.MoveSpeed = oLevel4Node.selectSingleNode("MoveSpeed").nodeTypedValue
            NewTpl.GravPow = oLevel4Node.selectSingleNode("GravPow").nodeTypedValue
            NewTpl.Inertia = oLevel4Node.selectSingleNode("Inertia").nodeTypedValue
            NewTpl.CollClass = oLevel4Node.selectSingleNode("CollClass").nodeTypedValue
            NewTpl.StateType = oLevel4Node.selectSingleNode("StateType").nodeTypedValue
            NewTpl.ControlType = oLevel4Node.selectSingleNode("ControlType").nodeTypedValue
            Set oLevel5Node = oLevel4Node.selectSingleNode("States")
            NewTpl.StateCount = oLevel5Node.selectSingleNode("@Count").Value
            For K = 0 To NewTpl.StateCount - 1
                Set oLevel6Node = oLevel5Node.selectSingleNode("State[@Index=" & CStr(K) & "]")
                If oLevel6Node.selectNodes("TSIdx").length > 0 Then
                    L = oLevel6Node.selectSingleNode("TSIdx").nodeTypedValue
                    If L >= 0 Then Set NewTpl.StateTilesetDef(K) = Prj.TileSetDef(L)
                End If
                Set oLevel7Node = oLevel6Node.selectSingleNode("Frames")
                For L = 0 To oLevel7Node.selectSingleNode("@Count").Value - 1
                    Set oLevel8Node = oLevel7Node.selectSingleNode("Frame[@Index=" & CStr(L) & "]")
                    NewTpl.AppendStateFrame K, oLevel8Node.nodeTypedValue
                Next
            Next
            NewMap.AddSpriteTemplate NewTpl
        Next
        Set NewTpl = Nothing
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("SpriteDefs")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("SpriteDef[@Index=" & CStr(J) & "]")
            Set NewSprDef = New SpriteDef
            With NewSprDef
                .Name = oLevel4Node.selectSingleNode("Name").nodeTypedValue
                Set .rLayer = NewMap.MapLayer(oLevel4Node.selectSingleNode("LayerName").nodeTypedValue)
                .Flags = oLevel4Node.selectSingleNode("Flags").nodeTypedValue
                Set .rPath = NewMap.Paths(oLevel4Node.selectSingleNode("Path").nodeTypedValue)
                If oLevel4Node.selectNodes("TemplateIndex").length > 0 Then
                    K = oLevel4Node.selectSingleNode("TemplateIndex").nodeTypedValue
                    If K >= 0 Then Set .Template = NewMap.SpriteTemplates(K)
                End If
            End With
            NewMap.AddSpriteDef NewSprDef
        Next
        Set NewSprDef = Nothing
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("CollDefs")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("CollDef[@Index=" & CStr(J) & "]")
            Set NewCollDef = New CollisionDef
            With NewCollDef
                .ClassA = oLevel4Node.selectSingleNode("ClassA").nodeTypedValue
                .ClassB = oLevel4Node.selectSingleNode("ClassB").nodeTypedValue
                .Flags = oLevel4Node.selectSingleNode("Flags").nodeTypedValue
                .SpecialFunction = oLevel4Node.selectSingleNode("SpecialFunction").nodeTypedValue
                .InvFlags = oLevel4Node.selectSingleNode("InvFlags").nodeTypedValue
                .InvItem = oLevel4Node.selectSingleNode("InvItem").nodeTypedValue
                .InvUseCount = oLevel4Node.selectSingleNode("InvUseCount").nodeTypedValue
                .Media = oLevel4Node.selectSingleNode("Media").nodeTypedValue
            End With
            NewMap.AddCollDef NewCollDef
        Next
        Set NewCollDef = Nothing

        Set oLevel3Node = oLevel2Node.selectSingleNode("CollClasses")
        If Not (oLevel3Node Is Nothing) Then
            For J = 0 To 15
                If oLevel3Node.selectNodes("CollClass[@Index=" & CStr(J) & "]").length > 0 Then
                    NewMap.CollClassName(J) = oLevel3Node.selectSingleNode("CollClass[@Index=" & CStr(J) & "]/Name").nodeTypedValue
                End If
            Next
        End If
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("Specials")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Special[@Index=" & CStr(J) & "]")
            Set NewSpecial = New SpecialFunction
            With NewSpecial
                .Name = oLevel4Node.selectSingleNode("Name").nodeTypedValue
                .LayerIndex = oLevel4Node.selectSingleNode("LayerIndex").nodeTypedValue
                .TileLeft = oLevel4Node.selectSingleNode("TileLeft").nodeTypedValue
                .TileTop = oLevel4Node.selectSingleNode("TileTop").nodeTypedValue
                .TileRight = oLevel4Node.selectSingleNode("TileRight").nodeTypedValue
                .TileBottom = oLevel4Node.selectSingleNode("TileBottom").nodeTypedValue
                .Flags = oLevel4Node.selectSingleNode("Flags").nodeTypedValue
                .FuncType = oLevel4Node.selectSingleNode("FuncType").nodeTypedValue
                .Value = oLevel4Node.selectSingleNode("Value").nodeTypedValue
                .SpriteName = oLevel4Node.selectSingleNode("SpriteName").nodeTypedValue
                .DestX = oLevel4Node.selectSingleNode("DestX").nodeTypedValue
                .DestY = oLevel4Node.selectSingleNode("DestY").nodeTypedValue
                .MediaName = oLevel4Node.selectSingleNode("MediaName").nodeTypedValue
                .InvItem = oLevel4Node.selectSingleNode("InvItem").nodeTypedValue
                .InvUseCount = oLevel4Node.selectSingleNode("InvUseCount").nodeTypedValue
            End With
            NewMap.AddSpecial NewSpecial
        Next
        Set NewSpecial = Nothing
        
        Set oLevel3Node = oLevel2Node.selectSingleNode("Interactions")
        For J = 0 To oLevel3Node.selectSingleNode("@Count").Value - 1
            Set oLevel4Node = oLevel3Node.selectSingleNode("Interaction[@Index=" & CStr(J) & "]")
            Set NewInteraction = New Interaction
            With NewInteraction
                If oLevel4Node.selectNodes("CategoryName").length > 0 Then
                    strKey = oLevel4Node.selectSingleNode("CategoryName").nodeTypedValue
                    strKey2 = oLevel4Node.selectSingleNode("TSName").nodeTypedValue
                    If Len(strKey) Then Set .TouchCategory = Prj.Groups(strKey, strKey2)
                End If
                .Flags = oLevel4Node.selectSingleNode("Flags").nodeTypedValue
                .Reaction = oLevel4Node.selectSingleNode("Reaction").nodeTypedValue
                .InvItem = oLevel4Node.selectSingleNode("InvItem").nodeTypedValue
                .ReplaceTile = oLevel4Node.selectSingleNode("ReplaceTile").nodeTypedValue
                .Media = oLevel4Node.selectSingleNode("Media").nodeTypedValue
            End With
            NewMap.AddInteraction NewInteraction
        Next
        Set NewInteraction = Nothing
    Next
    Set NewMap = Nothing

    Set oLevel8Node = Nothing
    Set oLevel7Node = Nothing
    Set oLevel6Node = Nothing
    Set oLevel5Node = Nothing
    Set oLevel4Node = Nothing
    Set oLevel3Node = Nothing
    Set oLevel2Node = Nothing
    Set oLevel1Node = Nothing
    Set oRoot = Nothing
    Set XMLDoc = Nothing
    XML2Prj = True
    Exit Function

LoadXMLErr:
    MsgBox "Error importing XML: " & Err.Description, vbExclamation
    XML2Prj = False
End Function
