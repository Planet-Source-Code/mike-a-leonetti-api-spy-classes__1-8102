VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4


Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Event LButtonDblClk()
Public Event LButtonDown()
Public Event LButtonUp()
Public Event RButtonDblClk()
Public Event RButtonDown()
Public Event RButtonUp()

Private WithEvents frmSysTray As PictureBox
Attribute frmSysTray.VB_VarHelpID = -1

Dim TrayIconData As NOTIFYICONDATA
Dim sToolTipText As String
Dim IconInTray As Boolean
Dim lRetValue As Long
Dim hIcon As Variant, lHandle As Long
Public Property Let ToolTipText(ByVal sToolTip As String)
    sToolTipText = sToolTip
    
    TrayIconData.szTip = sToolTipText & Chr(0)
    
    If IconInTray Then
        lRetValue = Shell_NotifyIcon(NIM_MODIFY, TrayIconData)
    End If
    
End Property
Public Property Get ToolTipText() As String
    ToolTipText = sToolTipText
End Property
Public Sub AddIconToSystemTray()
    
    lRetValue = Shell_NotifyIcon(NIM_ADD, TrayIconData)
    IconInTray = True
End Sub

Private Sub Class_Initialize()
    TrayIconData.cbSize = Len(TrayIconData)
    TrayIconData.uID = CLng(2)
    TrayIconData.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIconData.uCallbackMessage = WM_MOUSEMOVE
    TrayIconData.hIcon = CLng(0)
    TrayIconData.szTip = Chr(0)
    
    IconInTray = False
    hIcon = 0
    lHandle = 0
End Sub
Public Property Get Icon() As Variant
    Icon = hIcon
End Property
Public Property Let Icon(ByVal hNewIcon As Variant)

    hIcon = hNewIcon
    TrayIconData.hIcon = hIcon
    
    If IconInTray Then
        lRetValue = Shell_NotifyIcon(NIM_MODIFY, TrayIconData)
    End If
    
End Property
Public Sub RemoveIconFromSystemTray()
    
    TrayIconData.cbSize = Len(TrayIconData)
    TrayIconData.hWnd = frmSysTray.hWnd
    TrayIconData.uID = CLng(2)
    
    lRetValue = Shell_NotifyIcon(NIM_DELETE, TrayIconData)
End Sub
Public Property Set PPictureBox(ByVal NewBox As PictureBox)

    Set frmSysTray = NewBox
    frmSysTray.ScaleMode = 3
    frmSysTray.Visible = False
    TrayIconData.hWnd = frmSysTray.hWnd
    
    If IconInTray Then
        lRetValue = Shell_NotifyIcon(NIM_MODIFY, TrayIconData)
    End If
    
End Property
Public Property Get PPictureBox() As PictureBox
    Set PPictureBox = frmSysTray
End Property

Private Sub frmSysTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static Msg As Long, OldMsg As Long
    
    OldMsg = Msg
    Msg = X
   
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                RaiseEvent LButtonDblClk
            Case WM_LBUTTONDOWN:
                RaiseEvent LButtonDown
            Case WM_LBUTTONUP:
                RaiseEvent LButtonUp
            Case WM_RBUTTONDBLCLK:
                RaiseEvent RButtonDblClk
            Case WM_RBUTTONDOWN:
                RaiseEvent RButtonDown
            Case WM_RBUTTONUP:
                RaiseEvent RButtonUp
        End Select

End Sub
