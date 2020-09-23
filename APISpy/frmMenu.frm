VERSION 5.00
Begin VB.Form frmMenu 
   ClientHeight    =   1665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox pikSnapShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   450
      ScaleHeight     =   1260
      ScaleWidth      =   1590
      TabIndex        =   1
      Top             =   300
      Width           =   1590
   End
   Begin VB.PictureBox pikPageBG 
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   60
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   1200
      ScaleWidth      =   1710
      TabIndex        =   0
      Top             =   0
      Width           =   1710
   End
   Begin VB.Menu mnuSpecial 
      Caption         =   "Special"
      Begin VB.Menu mnuEnDis 
         Caption         =   "E&nable/Disable"
         Begin VB.Menu mnuDisable 
            Caption         =   "Disable Window"
         End
         Begin VB.Menu mnuEnable 
            Caption         =   "Enable Window"
         End
      End
      Begin VB.Menu mnuShowWindow1 
         Caption         =   "S&how Window"
         Begin VB.Menu mnuHideWindow 
            Caption         =   "Hide Window"
         End
         Begin VB.Menu mnuShowWindow 
            Caption         =   "Show Window"
         End
         Begin VB.Menu mnuMaximizeWindow 
            Caption         =   "Maximize Window"
         End
         Begin VB.Menu mnuMinimizeWindow 
            Caption         =   "Minimize Window"
         End
         Begin VB.Menu mnuRestore 
            Caption         =   "Restore Window"
         End
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Add &Message To Windows Queue"
         Begin VB.Menu mnuKey 
            Caption         =   "Key..."
            Begin VB.Menu mnuEnter 
               Caption         =   "Enter"
            End
         End
         Begin VB.Menu mnuClick 
            Caption         =   "Click"
            Begin VB.Menu mnuLBS 
               Caption         =   "Left Button (Single)"
            End
            Begin VB.Menu mnuLBD 
               Caption         =   "Left Button (Double)"
            End
            Begin VB.Menu mnuRBS 
               Caption         =   "Right Button (Single)"
            End
            Begin VB.Menu mnuRBD 
               Caption         =   "Right Button (Double)"
            End
         End
         Begin VB.Menu mnuCreate 
            Caption         =   "Create"
         End
         Begin VB.Menu mnuDoRePaint 
            Caption         =   "Paint"
         End
         Begin VB.Menu mnuDestroy 
            Caption         =   "Destroy"
         End
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
      End
      Begin VB.Menu mnuSet 
         Caption         =   "S&et..."
         Begin VB.Menu mnuNewParent 
            Caption         =   "Window's Parent"
         End
         Begin VB.Menu mnuSetWindowText 
            Caption         =   "Window Text"
         End
         Begin VB.Menu mnuSetControlText 
            Caption         =   "Control Text"
         End
         Begin VB.Menu mnuOTV 
            Caption         =   "OnTop Values"
            Begin VB.Menu mnuZOrder 
               Caption         =   "Z-Order"
               Begin VB.Menu mnuTopOfZOrder 
                  Caption         =   "Top Of Z-Order"
               End
               Begin VB.Menu mnuBottomOfZOrder 
                  Caption         =   "Bottom Of Z-Order"
               End
            End
            Begin VB.Menu mnuTAOT 
               Caption         =   "Always On Top"
            End
            Begin VB.Menu mnuRegularOT 
               Caption         =   "Standard"
            End
         End
      End
      Begin VB.Menu mnuBlock1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindWindow 
         Caption         =   "Find &Window..."
         Begin VB.Menu mnuFindWiz 
            Caption         =   "Using Wizzard"
         End
         Begin VB.Menu mnuBlock7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFindWindowByClassName 
            Caption         =   "By ClassName"
         End
         Begin VB.Menu mnuFindWindowByHandle 
            Caption         =   "By Handle"
         End
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Using Current Handle"
      End
      Begin VB.Menu mnuTakeShot 
         Caption         =   "Save Snapsh&ot..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save File..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuBlock3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Pr&ogram Preferences"
         Begin VB.Menu mnuAOT 
            Caption         =   "Always On Top"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSplash 
            Caption         =   "Show Splash Screen On Startup"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRepErr 
            Caption         =   "Report API Errors"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuBlock5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Begin VB.Menu mnuHelpWinQueue 
            Caption         =   "What Is ""Windows Queue?"""
         End
         Begin VB.Menu mnuHelpLockMode 
            Caption         =   "What Is Lock Mode?"
         End
         Begin VB.Menu mnuBlock4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "About"
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub SaveFile()
Dim iFreeFile As Integer

iFreeFile = FreeFile

If Dir(PROGRAM_SAVE_DIR_API, vbNormal) <> "" Then Kill PROGRAM_SAVE_DIR_API

    SavePicture pikPageBG.Picture, PROGRAM_SAVE_DIR_WPBG

Open PROGRAM_SAVE_DIR_API For Output As #iFreeFile
    Print #iFreeFile, "<html><head><title>" & VERSION_NAME & "'s API Spy Log</title></head><body background=" & Chr(34) & PROGRAM_SAVE_DIR_WPBG & Chr(34) & "><table width=395 border=0 cellspacing=0 cellpadding=0 height=50><tr>"
    Print #iFreeFile, "<td valign=top height=16 width=108 align=right><b></b></td><td valign=top height=16 width=215 align=center><b><font face=Arial><small><bold>---Main Window API---</bold></small></font></b></td><td valign=top height=16 width=72 align=left><b></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Handle:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtHMain.Text
    Print #iFreeFile, "</font></b></font></td></tr><tr> <td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Class Name:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtClassMain.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Window Text:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtTxtMain.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Control Text:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtMainCText.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Window State:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtMainState.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Position X:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtXMain.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Position Y:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtYMain.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Window Width:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtMainWidth.Text & "</font></b></font></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><font size=2><b><font face=Arial, Helvetica, sans-serif>Window Height:</font></b></font></td><td valign=top height=2 width=215 align=center><font size=2><b><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font size=2 face=Arial, Helvetica, sans-serif>-----</font><font face=Arial, Helvetica, sans-serif></font></b></font></td><td valign=top height=2 width=72 align=left><font size=2><b><font face=Arial, Helvetica, sans-serif>" & frmMain.txtMainHeight.Text & "</font></b></font></td></tr></table><br><br><table width=395 border=0 cellspacing=0 cellpadding=0 height=50>"
    
    Print #iFreeFile, "<tr><td valign=top height=16 width=108 align=right><b></b></td><td valign=top height=16 width=215 align=center><b><font face=Arial><small><bold>---Parent Window API---</bold></small></font></b></td><td valign=top height=16 width=72 align=left>&nbsp;</td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Handle:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtHParent.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Class Name:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtClassParent.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Window Text:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtTxtParent.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Control Text:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtParentCText.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Window State:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtParentState.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Position X:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtXParent.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Position Y:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtYParent.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Window Width:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtParentWidth.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Window Height:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtParentHeight.Text & "</font></b></font></td></tr></table><br><br><table width=395 border=0 cellspacing=0 cellpadding=0 height=50>"
    
    Print #iFreeFile, "<tr><td valign=top height=16 width=108 align=right><b></b></td><td valign=top height=16 width=215 align=center><b><font face=Arial><small><bold>---Color Values---</bold></small></font></b></td><td valign=top height=16 width=72 align=left>&nbsp;</td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Decimal:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtDec.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Hexadecimal:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtHex.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Red:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtR.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Green:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtG.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>Blue:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtB.Text & "</font></b></td></tr><tr>"
    Print #iFreeFile, "<td valign=top height=2 width=108 align=right><b><font size=2 face=Arial, Helvetica, sans-serif>BPP:</font></b></td><td valign=top height=2 width=215 align=center><b><font size=2 face=Arial, Helvetica, sans-serif>----------</font></b></td><td valign=top height=2 width=72 align=left><b><font size=2 face=Arial, Helvetica, sans-serif>" & frmMain.txtBPP.Text & "</font></b></td></tr></table></body></html>"
    
Close #iFreeFile

MsgBox "File saved to " & Chr(34) & PROGRAM_SAVE_DIR_API & Chr(34), vbInformation + vbSystemModal, "File Save Success"

End Sub


Private Sub mnuAbout_Click()
MsgBox "Special tribute to PhuryX13 for all his wonderful art!!!", vbSystemModal + vbInformation, "GO DAVE!!!"
frmAbout.Show vbModal
End Sub

Private Sub mnuAOT_Click()

If mnuAOT.Checked = True Then
    mnuAOT.Checked = False
    Call Frm_OnTop(frmMain, False)
    bMain_OnTop = False
Else
    mnuAOT.Checked = True
    Call Frm_OnTop(frmMain, True)
    bMain_OnTop = True
End If

End Sub
Private Sub mnuBottomOfZOrder_Click()
    lRetValue = SetWindowPos(hWin, HWND_BOTTOM, 0, 0, 0, 0, Flags)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuClose_Click()
    lRetValue = PostMessage(hWin, WM_CLOSE, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuCreate_Click()
lRetValue = PostMessage(hWin, WM_CREATE, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuDestroy_Click()
lRetValue = PostMessage(hWin, WM_DESTROY, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuDisable_Click()
lRetValue = EnableWindow(hWin, False)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuDoRePaint_Click()

lRetValue = PostMessage(hWin, WM_PAINT, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuEnable_Click()
lRetValue = EnableWindow(hWin, True)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuEnter_Click()
lRetValue = PostMessage(hWin, WM_CHAR, ByVal CLng(&HD), ByVal CLng(0))

End Sub
Private Sub mnuExit_Click()

frmMain.EndProgram
End Sub

Private Sub mnuFindWindowByClassName_Click()
Dim sClassName As String, hFW As Long

sClassName = Cus_InputBox(frmMain, "Enter Class Name", "Find Window By Class", "", CIB_ONTOP)

hFW = FindWindow(sClassName, CLng(0))

If hFW = 0 Then
    MsgBox "Window not found", vbInformation + vbSystemModal, "Error finding window"
Else
    frmMain.DoSpy (hFW)
End If

End Sub

Private Sub mnuFindWindowByHandle_Click()
Dim hRetVal As Long

hRetVal = Val(Cus_InputBox(frmMain, "Enter Handle Of Window", "Find Window By Handle", frmMain.hwnd, CIB_NOLETTERS Or CIB_ONTOP))

If IsWindow(hRetVal) = 0 Then
    MsgBox "Handle not associated with a window", vbInformation + vbSystemModal, "Error finding window"
Else
    Call frmMain.DoSpy(hRetVal)
End If

End Sub

Private Sub mnuFindWiz_Click()
Load frmShowAll
frmShowAll.Show
End Sub

Private Sub mnuHelpLockMode_Click()
frmHelpLockMode.Show vbModal

End Sub

Private Sub mnuHelpWinQueue_Click()
MsgBox "LOOK IT UP!  I CAN'T TELL YOU EVERYTHING!", vbSystemModal + vbInformation, "Buy a book on it!"

End Sub

Private Sub mnuHideWindow_Click()
lRetValue = ShowWindow(hWin, SW_HIDE)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuLBD_Click()

lRetValue = PostMessage(hWin, WM_LBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuLBS_Click()

lRetValue = PostMessage(hWin, WM_LBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
lRetValue = PostMessage(hWin, WM_LBUTTONUP, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuMaximizeWindow_Click()
lRetValue = ShowWindow(hWin, SW_MAXIMIZE)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuMinimizeWindow_Click()
lRetValue = ShowWindow(hWin, SW_MINIMIZE)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuNewParent_Click()
Dim hReturn As Long

hReturn = Val(Cus_InputBox(frmMain, "Enter Handle Of Window", "Set/Change Window Parent", "", CIB_NOLETTERS Or CIB_ONTOP))

If IsWindow(hReturn) = 0 Then
    MsgBox "Handle not associated with a window", vbInformation + vbSystemModal, "Error finding window"
Else
    lRetValue = SetParent(hWin, hReturn)
    If lRetValue = 0 Then Call ErrorDo
End If

End Sub

Private Sub mnuRBD_Click()

lRetValue = PostMessage(hWin, WM_RBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuRBS_Click()

lRetValue = PostMessage(hWin, WM_RBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
lRetValue = PostMessage(hWin, WM_RBUTTONUP, ByVal CLng(0), ByVal CLng(0))

End Sub

Private Sub mnuRefresh_Click()
Call frmMain.DoSpy(hWin)
End Sub

Private Sub mnuRegularOT_Click()
    lRetValue = SetWindowPos(hWin, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuRepErr_Click()

If mnuRepErr.Checked = True Then
    mnuRepErr.Checked = False
    bReportError = False
Else
    mnuRepErr.Checked = True
    bReportError = True
End If

End Sub

Private Sub mnuRestore_Click()
lRetValue = ShowWindow(hWin, SW_RESTORE)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuSave_Click()
SaveFile
End Sub

Private Sub mnuSavePref_Click()


End Sub

Private Sub mnuSetControlText_Click()
Dim sRet As String

NeedType = True
sRet = Cus_InputBox(frmMain, "Input new Control Text", "Change Control Text", sTxtMain, CIB_ONTOP)
NeedType = False

lRetValue = SendMessage(hWin, WM_SETTEXT, ByVal CLng(0), ByVal sRet)

End Sub

Private Sub mnuSetWindowText_Click()
Dim sRet As String

NeedType = True
sRet = Cus_InputBox(frmMain, "Input new Window Text", "Change Window Text", sTxtMain, CIB_ONTOP)
NeedType = False

lRetValue = SetWindowText(hWin, sRet)
    
    If lRetValue = 0 Then Call ErrorDo
    
End Sub

Private Sub mnuShowWindow_Click()
lRetValue = ShowWindow(hWin, SW_SHOW)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuSplash_Click()

If mnuSplash.Checked = True Then
    mnuSplash.Checked = False
    bShowSplash = False
Else
    mnuSplash.Checked = True
    bShowSplash = True
End If

End Sub

Private Sub mnuTakeShot_Click()
If hWin = 0 Then Exit Sub
Dim hWinDC As Long, lHeight As Long, lWidth As Long

lHeight = frmMain.txtMainHeight.Text
lWidth = frmMain.txtMainWidth.Text

pikSnapShot.Height = lHeight * Screen.TwipsPerPixelY
pikSnapShot.Width = lWidth * Screen.TwipsPerPixelY

hWinDC = GetWindowDC(hWin)
lRetValue = BitBlt(pikSnapShot.hDC, 0, 0, lWidth, lHeight, hWinDC, 0, 0, SRCCOPY)
lRetValue = ReleaseDC(hWin, hWinDC)
Call SavePicture(pikSnapShot.Image, PROGRAM_SAVE_DIR_SNAP)
pikSnapShot.Height = 1
pikSnapShot.Width = 1
pikSnapShot.Cls

MsgBox "File saved to " & Chr(34) & PROGRAM_SAVE_DIR_SNAP & Chr(34), vbInformation + vbSystemModal, "File Save Success"

End Sub

Private Sub mnuTAOT_Click()
    lRetValue = SetWindowPos(hWin, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

Private Sub mnuTopOfZOrder_Click()
    lRetValue = SetWindowPos(hWin, HWND_TOP, 0, 0, 0, 0, Flags)
    
    If lRetValue = 0 Then Call ErrorDo
End Sub

