Attribute VB_Name = "InFeStEdT"
Option Explicit

Enum EnumCusInputBox
    CIB_NOLETTERS = &H2
    CIB_ONTOP = &H4
End Enum

'Preferences
Public bMain_OnTop As Boolean
Public bReportError As Boolean
Public bShowSplash As Boolean

Public PROGRAM_SAVE_DIR_PREF As String
Public PROGRAM_SAVE_DIR_WPBG As String

Public PROGRAM_SAVE_DIR_API As String
Public PROGRAM_SAVE_DIR_SNAP As String

Public Const VERSION_NAME = "InFeStEd"

Public MainWins As Collection, ChildWins As Collection
Public CurPos As Point_API, hWin As Long, sRetValue As String, bNoLetters As Boolean
Public hParent As Long, lParentTextLen As Long, lMainTextLen As Long
Public sClassNameMain As String, sClassNameParent As String
Public sTxtMain As String, sTxtParent As String, MainDC As Long
Public rectMain As Rect, rectParent As Rect, CurColor As COLORRGB
Public lCurColor As Long, ParentDC As Long, lRetValue As Long
Public NeedType As Boolean, byMainBPP As Byte
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueNumber As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Const SW_SHOWNORMAL = 5

Public Const SRCCOPY = &HCC0020

Public Const BITSPIXEL = 12

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_BOTTOM = 1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const WM_CLOSE = &H10
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = 3
Public Const WM_SIZE = 5
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_PAINT = &HF

Const ERROR = 0
Const NULLREGION = 1
Const SIMPLEREGION = 2
Const COMPLEXREGION = 3
Const RGN_AND = 1
Const RGN_OR = 2
Const RGN_XOR = 3
Const RGN_DIFF = 4
Const RGN_COPY = 5

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

Type Point_API
    X As Long
    Y As Long
End Type

Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As Point_API) As Long
Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As Rect) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Function RGB_GetRGB(ByVal CVal As Long) As COLORRGB
Dim TempColor As COLORRGB
    
    TempColor.Blue = Int(CVal / 65536)
    TempColor.Green = Int((CVal - (65536 * TempColor.Blue)) / 256)
    TempColor.Red = CVal - (65536 * TempColor.Blue + 256 * TempColor.Green)

RGB_GetRGB = TempColor

  
End Function
Public Function Sys_GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String, GetFromINI As String
      strBuffer = String(750, Chr(0))
   Key = LCase(Key)
   Sys_GetFromINI = Left(strBuffer, GetPrivateProfileString(Section, ByVal Key, "", strBuffer, Len(strBuffer), Directory))
End Function
Public Function Sys_iGetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String, GetFromINI As String
      strBuffer = String(750, Chr(0))
   Key = LCase(Key)
   Sys_iGetFromINI = Left(strBuffer, GetPrivateProfileString(Section, ByVal Key, 0, strBuffer, Len(strBuffer), Directory))
End Function
Sub Frm_FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, ByVal CLng(2), ByVal CLng(0))
End Sub
Public Function Sys_WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)

    Key = UCase(Key)
    
    Sys_WriteToINI = WritePrivateProfileString(Section, Key, KeyValue, Directory)
End Function
Sub Frm_OnTop(TheForm As Form, OnTop As Boolean)

Dim SetWinOnTop As Long

    If OnTop Then
        SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If
    
End Sub
Sub ErrorDo()
If bReportError Then
    MsgBox "Error initializing task, most likely not in my program!" & vbNewLine & vbNewLine & "Refresh the window to see if it still exists.", vbInformation + vbSystemModal, "Error!"
End If
End Sub
Sub Tmr_Delay(lSeconds As Long)
Dim Time1 As Double, Time2 As Double

Time1 = Timer
Time2 = Timer + lSeconds

While Time1 < Time2
    Time1 = Timer
    DoEvents
Wend

End Sub
Function NTW_GetUserName() As String
    Dim lpBuffer As String
    
    lpBuffer = Space(255)
    lRetValue = GetUserName(lpBuffer, 255)
    NTW_GetUserName = StripNulls(lpBuffer)
    
End Function

Private Function StripNulls(S As String) As String
    Dim I As Integer
    StripNulls = S


    If Len(S) Then
        I = InStr(S, Chr(0))
        If I Then StripNulls = Left(S, I - 1)
    End If
End Function
Function Cus_InputBox(frmOwner As Form, sPrompt As String, sTitle As String, sDefault As String, lFlags As EnumCusInputBox) As String

    Load frmInputBox
    frmInputBox.Caption = sTitle
    frmInputBox.txtMain.Text = sDefault
    frmInputBox.lblText.Caption = sPrompt
    
    If lFlags And CIB_NOLETTERS Then
        bNoLetters = True
    Else
        bNoLetters = False
    End If
    
    If lFlags And CIB_ONTOP Then
        Call Frm_OnTop(frmInputBox, True)
    End If
    
    frmInputBox.Show vbModal, frmOwner
    Cus_InputBox = sRetValue
End Function

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    ChildWins.Add hWnd
    
    EnumChildProc = 1
End Function
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    MainWins.Add hWnd
  
    EnumWindowsProc = 1
End Function

Sub NET_JumpToWebsite(ByVal sUrl As String)

Call ShellExecute(CLng(0), vbNullString, sUrl, vbNullString, vbNullString, vbNormalFocus)

End Sub
Public Function ShapeFormByColor(TheForm As Form, TColor As Variant) As Boolean
Dim lX As Long, lY As Long, hTempRegion As Long, FormWidth As Long, FormHeight As Long
Dim hCombinedRgn As Long, TempColor As Variant

'This function was from planet-source-code.com, I wasn't smart enough to write it :(
FormWidth = TheForm.Width / 15
FormHeight = TheForm.Height / 15

hCombinedRgn = CreateRectRgn(0, 0, FormWidth, FormHeight)

While lY <= FormHeight
    While lX <= FormWidth
    
        TempColor = GetPixel(TheForm.hDC, lX, lY)
        
        If TempColor = TColor Then
            hTempRegion = CreateRectRgn(lX, lY, lX + 1, lY + 1)
            Call CombineRgn(hCombinedRgn, hCombinedRgn, hTempRegion, RGN_DIFF)
            Call DeleteObject(hTempRegion)
        End If
        lX = lX + 1
    Wend
        lY = lY + 1
        lX = 0
Wend

ShapeFormByColor = SetWindowRgn(TheForm.hWnd, hCombinedRgn, True)
Call DeleteObject(hCombinedRgn)

End Function
