VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4200
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Click Here"
      Height          =   195
      Left            =   5070
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim frmMainX As Long, frmMainY As Long, bFileFound As Boolean, sUser As String
Dim sSysDir As String, sWinPath As String

On Error Resume Next

Load frmMenu
Load frmMain

frmMain.BackColor = RGB(255, 0, 150)
ShapeFormByColor frmMain, RGB(255, 0, 150)

sSysDir = Space(255)
lRetValue = GetSystemDirectory(sSysDir, 255)
sSysDir = Left(sSysDir, lRetValue)

sWinPath = Space(255)
lRetValue = GetWindowsDirectory(sWinPath, 255)
sWinPath = Left(sWinPath, lRetValue)

If Right(sSysDir, 1) = "\" Then
    PROGRAM_SAVE_DIR_PREF = sSysDir & "\APISpySave.API"
    PROGRAM_SAVE_DIR_WPBG = sSysDir & "\WPBGSpy.bmp"
Else
    PROGRAM_SAVE_DIR_PREF = sSysDir & "\APISpySave.API"
    PROGRAM_SAVE_DIR_WPBG = sSysDir & "\WPBGSpy.bmp"
End If

If Right(sWinPath, 1) = "\" Then
    PROGRAM_SAVE_DIR_API = sWinPath & "Desktop\API_Log.html"
    PROGRAM_SAVE_DIR_SNAP = sWinPath & "Desktop\API_SnapShot.bmp"
Else
    PROGRAM_SAVE_DIR_API = sWinPath & "\Desktop\API_Log.API.html"
    PROGRAM_SAVE_DIR_SNAP = sWinPath & "\Desktop\API_SnapShot.bmp"
End If

sUser = NTW_GetUserName & " - " & Date & " (" & Time & ")"

If Not (Dir(PROGRAM_SAVE_DIR_PREF, vbNormal)) <> "" Then
    bMain_OnTop = True
    bReportError = True
    bShowSplash = True
    frmSplash.Show
    Call Frm_OnTop(frmSplash, True)
    frmSplash.Caption = "Loading " & VERSION_NAME & " 's API Spy (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")..."
    lRetValue = Sys_WriteToINI("userlog", "user1", sUser, PROGRAM_SAVE_DIR_PREF)
    lRetValue = Sys_WriteToINI("userlog", "userentries", 1, PROGRAM_SAVE_DIR_PREF)
Else
    
    bShowSplash = Sys_GetFromINI("main", "splash", PROGRAM_SAVE_DIR_PREF)
    frmMenu.mnuSplash.Checked = bShowSplash
    If bShowSplash Then
        frmSplash.Show
        Call Frm_OnTop(frmSplash, True)
        frmSplash.Caption = "Loading " & VERSION_NAME & " 's API Spy (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")..."
    End If
    
    bMain_OnTop = Sys_iGetFromINI("main", "aot", PROGRAM_SAVE_DIR_PREF)
    bReportError = Sys_iGetFromINI("main", "errrpt", PROGRAM_SAVE_DIR_PREF)
    frmMainX = Sys_iGetFromINI("pos", "x", PROGRAM_SAVE_DIR_PREF)
    frmMainY = Sys_iGetFromINI("pos", "y", PROGRAM_SAVE_DIR_PREF)
    frmMenu.mnuAOT.Checked = bMain_OnTop
    frmMenu.mnuRepErr.Checked = bReportError
    bFileFound = True
    
    Dim iUserEntry As Integer
    
    iUserEntry = Sys_iGetFromINI("userlog", "userentries", PROGRAM_SAVE_DIR_PREF)
    
    If iUserEntry = 0 Then
        lRetValue = Sys_WriteToINI("userlog", "user1", sUser, PROGRAM_SAVE_DIR_PREF)
        lRetValue = Sys_WriteToINI("userlog", "userentries", 1, PROGRAM_SAVE_DIR_PREF)
    Else
        lRetValue = Sys_WriteToINI("userlog", "user" & (iUserEntry + 1), sUser, PROGRAM_SAVE_DIR_PREF)
        lRetValue = Sys_WriteToINI("userlog", "userentries", iUserEntry + 1, PROGRAM_SAVE_DIR_PREF)
    End If
End If
    

frmMain.Caption = VERSION_NAME & "'s API Spy (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")"

If bShowSplash Then
    Tmr_Delay 4
End If

Unload frmSplash

frmMain.Show

Call Frm_OnTop(frmMain, bMain_OnTop)

frmMain.AddIcon

If bFileFound Then
    frmMain.Move frmMainX, frmMainY
Else
    frmMain.Move 0, 0
End If
End Sub

Private Sub Label1_Click()
MsgBox "Mike, make me a sub to let images fade into each other on mouseover and then to fade back on mouseout. This'll add a little size to the program but it'll look cool. Also possibly maybe you can make one to make the background of textboxes transparent" & vbNewLine & vbNewLine & "-Dave", vbOKOnly & vbSystemModal, "ToDo List"
End Sub
