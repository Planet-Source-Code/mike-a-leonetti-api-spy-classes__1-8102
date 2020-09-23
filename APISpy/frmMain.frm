VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4890
   ClientLeft      =   2985
   ClientTop       =   0
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   4890
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pikForSYSTRAY 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6480
      ScaleHeight     =   360
      ScaleWidth      =   375
      TabIndex        =   35
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtHex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4140
      Width           =   795
   End
   Begin VB.TextBox txtDec 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3900
      Width           =   795
   End
   Begin VB.TextBox txtB 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4380
      Width           =   435
   End
   Begin VB.TextBox txtG 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4140
      Width           =   435
   End
   Begin VB.TextBox txtR 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3900
      Width           =   435
   End
   Begin VB.TextBox txtXMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtXParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtYMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2910
      Width           =   1455
   End
   Begin VB.TextBox txtBPP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3330
      Width           =   1455
   End
   Begin VB.TextBox txtMainState 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2610
      Width           =   1305
   End
   Begin VB.TextBox txtMainHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   1305
   End
   Begin VB.TextBox txtMainWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2970
      Width           =   1305
   End
   Begin VB.TextBox txtMainCText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2160
      Width           =   3285
   End
   Begin VB.TextBox txtTxtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   3165
   End
   Begin VB.TextBox txtClassMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1350
      Width           =   3255
   End
   Begin VB.TextBox txtHMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   3855
   End
   Begin VB.OptionButton optMain 
      BackColor       =   &H00404040&
      Height          =   195
      Left            =   2010
      TabIndex        =   7
      Top             =   4140
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton optParent 
      BackColor       =   &H00404040&
      Height          =   195
      Left            =   2010
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   3900
      Width           =   195
   End
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   6300
      Top             =   1350
   End
   Begin VB.CheckBox chkLock 
      BackColor       =   &H00000000&
      Height          =   195
      Left            =   2010
      TabIndex        =   4
      Top             =   4380
      Width           =   195
   End
   Begin VB.PictureBox pikColor 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2400
      ScaleHeight     =   705
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   3900
      Width           =   735
   End
   Begin VB.TextBox txtHParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtClassParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1350
      Width           =   3285
   End
   Begin VB.TextBox txtTxtParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1800
      Width           =   3165
   End
   Begin VB.TextBox txtParentCText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   3285
   End
   Begin VB.TextBox txtYParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2910
      Width           =   1455
   End
   Begin VB.TextBox txtParentHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3360
      Width           =   1305
   End
   Begin VB.TextBox txtParentWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2970
      Width           =   1305
   End
   Begin VB.TextBox txtParentState 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2610
      Width           =   1305
   End
   Begin VB.Image imgStop 
      Height          =   915
      Left            =   6960
      Picture         =   "frmMain.frx":A76F
      Top             =   2700
      Width           =   420
   End
   Begin VB.Image imgStart 
      Height          =   915
      Left            =   6420
      Picture         =   "frmMain.frx":ADB0
      Top             =   2010
      Width           =   345
   End
   Begin VB.Image imgStartStop 
      Height          =   915
      Left            =   240
      Top             =   1770
      Width           =   345
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4410
      TabIndex        =   1
      Top             =   150
      Width           =   195
   End
   Begin VB.Label lblMinimize 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4230
      TabIndex        =   2
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblMenu 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   540
      TabIndex        =   3
      Top             =   270
      Width           =   810
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   4380
   End
   Begin VB.Label lblParent 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1290
      TabIndex        =   8
      Top             =   4110
      Width           =   525
   End
   Begin VB.Shape shpLock 
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Left            =   1290
      Shape           =   2  'Oval
      Top             =   4380
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblLock 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   4380
      Width           =   660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bEnabled As Boolean
Private WithEvents SystemTray As clsSystemTray
Attribute SystemTray.VB_VarHelpID = -1
Private Sub cmdExit_Click()
End
End Sub

Private Sub chkLock_Click()
If chkLock.Value = vbChecked Then
    imgStartStop.Enabled = False

        
Else

    imgStartStop.Enabled = True
End If
End Sub

Private Sub cmdMin_Click()
frmMain.WindowState = vbMinimized
End Sub
Sub DoSpy(Optional phWnd As Long = 0)

On Error Resume Next

    'Allocate space
    sClassNameMain = Space(255)
    sClassNameParent = Space(255)
    
    'Get Cursor position
    lRetValue = GetCursorPos(CurPos)

    If phWnd = 0 Then
        'Get Main Handle
        hWin = WindowFromPoint(CurPos.X, CurPos.Y)
    Else
        hWin = phWnd
    End If
    
        txtHMain.Text = hWin
        
    'Get Main Class Name
    lRetValue = GetClassName(hWin, sClassNameMain, 255)
    sClassNameMain = Left(sClassNameMain, lRetValue)
    txtClassMain.Text = sClassNameMain
    
    'Get Main Text
    lMainTextLen = GetWindowTextLength(hWin) + 1
    sTxtMain = Space(lMainTextLen)
    lRetValue = GetWindowText(hWin, sTxtMain, lMainTextLen)
    sTxtMain = Left(sTxtMain, lMainTextLen)
    txtTxtMain.Text = sTxtMain
    
    lRetValue = SendMessage(hWin, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
    sTxtMain = Space(lRetValue)
    lRetValue = SendMessage(hWin, WM_GETTEXT, ByVal lRetValue, ByVal sTxtMain)
    txtMainCText.Text = sTxtMain
    
    'Get Main Rect
    lRetValue = GetWindowRect(hWin, rectMain)
    txtXMain.Text = rectMain.Left
    txtYMain.Text = rectMain.Top
    txtMainHeight.Text = rectMain.Bottom - rectMain.Top
    txtMainWidth.Text = rectMain.Right - rectMain.Left
    
    'Get Main State
    If (Not IsIconic(hWin)) And (Not IsZoomed(hWin)) Then txtMainState.Text = "General"
    If IsIconic(hWin) Then txtMainState.Text = "Minimized"
    If IsZoomed(hWin) Then txtMainState.Text = "Maximixed"
    
    'Get Parent Handle
    hParent = GetParent(hWin)
    If hParent <> 0 Then
        txtHParent.Text = hParent
        
        'Get Parent Class Name
        lRetValue = GetClassName(hParent, sClassNameParent, 255)
        sClassNameParent = Left(sClassNameParent, lRetValue)
        txtClassParent.Text = sClassNameParent
        
        'Get Parent Text
        lParentTextLen = GetWindowTextLength(hParent) + 1
        sTxtParent = Space(lParentTextLen)
        lRetValue = GetWindowText(hParent, sTxtParent, lParentTextLen)
        sTxtParent = Left(sTxtParent, lParentTextLen)
        txtTxtParent.Text = sTxtParent
        
            
        lRetValue = SendMessage(hParent, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
        sTxtParent = Space(lRetValue)
        lRetValue = SendMessage(hParent, WM_GETTEXT, ByVal lRetValue, ByVal sTxtParent)
        txtParentCText.Text = sTxtParent
        
        'Get Parent Rect
        lRetValue = GetWindowRect(hParent, rectParent)
        txtXParent.Text = rectParent.Left
        txtYParent.Text = rectParent.Top
        txtParentHeight.Text = rectParent.Bottom - rectParent.Top
        txtParentWidth.Text = rectParent.Right - rectParent.Left
        
        'Get Parent State
        If (Not IsIconic(hParent)) And (Not IsZoomed(hParent)) Then txtParentState.Text = "General"
        If IsIconic(hParent) Then txtParentState.Text = "Minimized"
        If IsZoomed(hParent) Then txtParentState.Text = "Maximixed"
    Else
        txtHParent.Text = "Window Does Not Have A Parent"
        txtClassParent.Text = "--"
        txtTxtParent.Text = "--"
        txtParentCText.Text = "--"
        txtXParent.Text = "--"
        txtYParent.Text = "--"
        txtParentHeight.Text = "--"
        txtParentWidth.Text = "--"
        txtParentState.Text = "--"
    End If
    
    'Get Colors
    MainDC = GetDC(0)
    lCurColor = GetPixel(MainDC, CurPos.X, CurPos.Y)
    CurColor = RGB_GetRGB(lCurColor)
    txtR.Text = CurColor.Red
    txtG.Text = CurColor.Green
    txtB.Text = CurColor.Blue
    pikColor.BackColor = lCurColor
    txtHex.Text = Hex(lCurColor)
    txtDec.Text = lCurColor
    lRetValue = ReleaseDC(0, MainDC)
    
    MainDC = GetWindowDC(hWin)
    byMainBPP = GetDeviceCaps(MainDC, BITSPIXEL)
    lRetValue = ReleaseDC(hWin, MainDC)
    
    txtBPP.Text = byMainBPP
End Sub

Sub AddIcon()
Set SystemTray = New clsSystemTray

SystemTray.Icon = frmMain.Icon
Set SystemTray.PPictureBox = pikForSYSTRAY
SystemTray.ToolTipText = VERSION_NAME & "'s API Spy (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")"

SystemTray.AddIconToSystemTray
End Sub
Sub EndProgram()

Dim sTempAOT As String, sTempERR As String, sTempSplash As String

If bMain_OnTop Then
    sTempAOT = 1
Else
    sTempAOT = 0
End If

If bReportError Then
    sTempERR = 1
Else
    sTempERR = 0
End If

If bShowSplash Then
    sTempSplash = 1
Else
    sTempSplash = 0
End If

lRetValue = Sys_WriteToINI("main", "aot", sTempAOT, PROGRAM_SAVE_DIR_PREF)
lRetValue = Sys_WriteToINI("main", "errrpt", sTempERR, PROGRAM_SAVE_DIR_PREF)
lRetValue = Sys_WriteToINI("main", "splash", sTempSplash, PROGRAM_SAVE_DIR_PREF)
lRetValue = Sys_WriteToINI("pos", "x", frmMain.Left, PROGRAM_SAVE_DIR_PREF)
lRetValue = Sys_WriteToINI("pos", "y", frmMain.Top, PROGRAM_SAVE_DIR_PREF)

SystemTray.RemoveIconFromSystemTray

On Error Resume Next

Unload frmMain
Unload frmInputBox
Unload frmHelpLockMode
Unload frmAbout
Unload frmShowAll

End

End Sub

Private Sub cmdmnuDisable_Click()

End Sub

Private Sub Form_Load()

imgStartStop.Picture = imgStart.Picture

End Sub

Private Sub Form_Unload(Cancel As Integer)
EndProgram
End Sub
Private Sub Image1_Click()

End Sub

Private Sub imgStartStop_Click()
If imgStartStop.Picture = imgStart.Picture Then
bEnabled = True

chkLock.Enabled = False
lblLock.Enabled = False
shpLock.Visible = True
imgStartStop.Picture = imgStop.Picture
Else
bEnabled = False

chkLock.Enabled = True
lblLock.Enabled = True
shpLock.Visible = False
imgStartStop.Picture = imgStart.Picture
End If
End Sub

Private Sub lblExit_Click()
EndProgram
End Sub

Private Sub lblLock_Click()
If chkLock.Value = vbChecked Then
    chkLock.Value = vbUnchecked
Else
    chkLock.Value = vbChecked
End If

chkLock_Click
End Sub

Private Sub lblMain_Click()

optMain.Value = True
optMain_Click


End Sub

Private Sub lblMenu_Click()
frmMain.PopupMenu frmMenu.mnuSpecial, 0, lblMenu.Left, lblMenu.Top + lblMenu.Height
End Sub

Private Sub lblMinimize_Click()
frmMain.Visible = False
End Sub



Private Sub lblParent_Click()

optParent.Value = True
optParent_Click

End Sub

Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Frm_FormDrag frmMain

End Sub

Private Sub optMain_Click()

txtHMain.Visible = True
txtClassMain.Visible = True
txtTxtMain.Visible = True
txtMainCText.Visible = True
txtMainState.Visible = True
txtMainWidth.Visible = True
txtYMain.Visible = True
txtXMain.Visible = True
txtMainHeight.Visible = True

txtHParent.Visible = False
txtClassParent.Visible = False
txtTxtParent.Visible = False
txtParentCText.Visible = False
txtParentState.Visible = False
txtParentWidth.Visible = False
txtYParent.Visible = False
txtXParent.Visible = False
txtParentHeight.Visible = False

End Sub

Private Sub optParent_Click()

txtHParent.Visible = True
txtClassParent.Visible = True
txtTxtParent.Visible = True
txtParentCText.Visible = True
txtParentState.Visible = True
txtParentWidth.Visible = True
txtYParent.Visible = True
txtXParent.Visible = True
txtParentHeight.Visible = True

txtHMain.Visible = False
txtClassMain.Visible = False
txtTxtMain.Visible = False
txtMainCText.Visible = False
txtMainState.Visible = False
txtMainWidth.Visible = False
txtYMain.Visible = False
txtXMain.Visible = False
txtMainHeight.Visible = False

End Sub

Private Sub SystemTray_LButtonDblClk()
If frmMain.Visible = True Then Exit Sub

frmMain.Visible = True
End Sub

Private Sub SystemTray_RButtonUp()
frmMain.PopupMenu frmMenu.mnuSpecial
End Sub

Private Sub tmrMain_Timer()
If chkLock.Value = vbChecked And Not NeedType Then
    
    If GetAsyncKeyState(vbKeyL) Then
        lRetValue = SetForegroundWindow(frmMain.hWnd)
        DoSpy
    End If

Else

    If bEnabled = True Then
        
        DoSpy
    
    End If
    
End If

End Sub

