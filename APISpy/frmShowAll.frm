VERSION 5.00
Begin VB.Form frmShowAll 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Window In System"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmShowAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pikFramed2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   3900
      ScaleHeight     =   2880
      ScaleWidth      =   3165
      TabIndex        =   11
      Top             =   285
      Width           =   3165
      Begin VB.PictureBox pikItems1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   0
         ScaleHeight     =   2715
         ScaleWidth      =   2865
         TabIndex        =   13
         Top             =   0
         Width           =   2925
         Begin VB.Label lblItemC 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Item1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.VScrollBar vsrChild 
         Height          =   2790
         Left            =   2970
         TabIndex        =   12
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.PictureBox pikFramed1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   210
      ScaleHeight     =   2880
      ScaleWidth      =   3465
      TabIndex        =   7
      Top             =   270
      Width           =   3465
      Begin VB.PictureBox pikItems 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2820
         Left            =   0
         ScaleHeight     =   2760
         ScaleWidth      =   3180
         TabIndex        =   9
         Top             =   0
         Width           =   3240
         Begin VB.Label lblItemP 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Item1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   -15
            TabIndex        =   10
            Top             =   -15
            Width           =   1920
         End
      End
      Begin VB.VScrollBar vsrParent 
         Height          =   2820
         Left            =   3285
         TabIndex        =   8
         Top             =   -15
         Width           =   165
      End
   End
   Begin VB.CheckBox chkVisible 
      BackColor       =   &H00000000&
      Caption         =   "Visible Windows Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4770
      TabIndex        =   4
      Top             =   3510
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   30
      Width           =   1425
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   180
      Y2              =   3375
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   3750
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   210
      X2              =   150
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3750
      X2              =   3750
      Y1              =   210
      Y2              =   3375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1710
      X2              =   3750
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3840
      X2              =   7140
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   210
      Y2              =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3930
      X2              =   3840
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   7140
      X2              =   7140
      Y1              =   210
      Y2              =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   5430
      X2              =   7140
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Child Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label lblLockWindow 
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   3765
      Width           =   1305
   End
   Begin VB.Label lblRefresh 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   3465
      Width           =   765
   End
   Begin VB.Label lblShowHandle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1500
      TabIndex        =   1
      ToolTipText     =   "Handle"
      Top             =   3765
      Width           =   2790
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Number"
      Top             =   3495
      Width           =   3390
   End
End
Attribute VB_Name = "frmShowAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSelIndex(1 To 2) As Long, hSpyWnd As Long

Private Sub chkVisible_Click()
Dim iRet As Integer

iRet = MsgBox("No changes will take place until you refresh.  Refresh now?", vbInformation + vbYesNo, "Refresh now?")

If iRet = vbYes Then
    lblRefresh_Click
End If
End Sub

Private Sub chkVisible_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chkVisible.ForeColor = &HC0C0C0
End Sub

Private Sub Form_Load()
Call Frm_OnTop(frmShowAll, True)
Call lblRefresh_Click

End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HC0C0C0
lblLockWindow.ForeColor = &HFFFFFF
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF
chkVisible.ForeColor = &HFFFFFF

End Sub
Private Sub lblGetChildren_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF

End Sub
Private Sub lblItemC_Click(Index As Integer)
If lblItemC(Index).Tag = "NoGo" Then Exit Sub

If lSelIndex(1) <> -1 Then
    If lSelIndex(2) = 2 Then
        lblItemC(lSelIndex(1)).BackColor = vbBlack
        lblItemC(lSelIndex(1)).ForeColor = vbWhite
    Else
        lblItemP(lSelIndex(1)).BackColor = vbBlack
        lblItemP(lSelIndex(1)).ForeColor = vbWhite
    End If
End If
    lSelIndex(1) = Index
    lSelIndex(2) = 2
    lblItemC(lSelIndex(1)).BackColor = vbGreen
    lblItemC(lSelIndex(1)).ForeColor = vbBlack
    lblShowHandle.Caption = "Item Handle:" & lblItemC(lSelIndex(1)).Tag
    lblNum.Caption = "Item Number:  " & lSelIndex(1) + 1
    hSpyWnd = lblItemC(lSelIndex(1)).Tag
End Sub

Private Sub lblItemC_DblClick(Index As Integer)
frmMain.DoSpy hSpyWnd
End Sub

Private Sub lblItemP_Click(Index As Integer)
If lSelIndex(1) <> -1 Then
    If lSelIndex(2) = 2 Then
        lblItemC(lSelIndex(1)).BackColor = vbBlack
        lblItemC(lSelIndex(1)).ForeColor = vbWhite
    Else
        lblItemP(lSelIndex(1)).BackColor = vbBlack
        lblItemP(lSelIndex(1)).ForeColor = vbWhite
    End If
End If
    lSelIndex(1) = Index
    lSelIndex(2) = 1
    lblItemP(lSelIndex(1)).BackColor = vbGreen
    lblItemP(lSelIndex(1)).ForeColor = vbBlack
    lblShowHandle.Caption = "Item Handle:" & lblItemP(lSelIndex(1)).Tag
    lblNum.Caption = "Item Number:  " & lSelIndex(1) + 1
    
    hSpyWnd = lblItemP(lSelIndex(1)).Tag
    
    GetChildrenOfSel
End Sub


Private Sub lblLockChildren_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF

End Sub

Private Sub lblItemP_DblClick(Index As Integer)
frmMain.DoSpy hSpyWnd
End Sub

Private Sub lblLockWindow_Click()
frmMain.DoSpy hSpyWnd
End Sub
Private Sub lblLockWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HC0C0C0

End Sub

Private Sub lblNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF

End Sub

Private Sub lblRefresh_Click()
Dim L As Long, sClassRet As String

ClearListMain
Set MainWins = New Collection

lRetValue = EnumWindows(AddressOf EnumWindowsProc, 0)

Dim IsVisible As Boolean

For L = 0 To MainWins.Count - 1

    IsVisible = IsWindowVisible(MainWins.Item(L + 1))
    
    If chkVisible.Value = vbUnchecked Then IsVisible = True
    
        sClassRet = Space(255)
        lRetValue = GetClassName(MainWins.Item(L + 1), sClassRet, 255)
        sClassRet = Left(sClassRet, lRetValue)
    
        If Not (L = 0) Then
            Load lblItemP(L)
        
        
            With lblItemP(L)
                .Left = 1
                If IsVisible Then
                    .Top = (lblItemP(L - 1).Height + lblItemP(L - 1).Top)
                Else
                    .Top = lblItemP(L - 1).Top
                End If
                .Width = pikItems.Width - 1
                .Height = 270
                .Tag = MainWins.Item(L + 1)
                .Caption = sClassRet
                .Visible = IsVisible
            End With
        Else
            With lblItemP(0)
                .Left = 1
                If IsVisible Then
                    .Top = 0
                Else
                    .Top = -270
                End If
                .Width = pikItems.Width - 1
                .Height = 270
                .Tag = MainWins.Item(1)
                .Caption = sClassRet
            End With
        End If
Next L


Dim iCount As Integer
iCount = 0

For L = 0 To lblItemP.ubound
    If lblItemP(L).Visible Then iCount = iCount + 1
Next L

pikItems.Height = (lblItemP(0).Height * (iCount)) + 45
vsrParent.Max = (pikFramed1.Height - pikItems.Height + 45) / (lblItemP(0).Height)
vsrParent.Visible = pikItems.Height > pikFramed1.Height
pikFramed1.Height = (lblItemP(0).Height * 11) + 45
vsrParent.Height = pikFramed1.Height
pikItems.Top = 0

lSelIndex(1) = 0
lSelIndex(2) = 1
GetChildrenOfSel

End Sub

Private Sub lblRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HC0C0C0
lblLockWindow.ForeColor = &HFFFFFF

End Sub

Private Sub lblShowHandle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF
End Sub
Private Sub lstChilds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF

End Sub


Private Sub lstWins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.ForeColor = &HFFFFFF
lblLockWindow.ForeColor = &HFFFFFF

End Sub
Sub ClearListMain()
Dim L As Long

On Error Resume Next

For L = 1 To lblItemP.ubound + 1
        Unload lblItemP(L)
Next L

End Sub

Private Sub vsrChild_Change()
pikItems1.Top = vsrChild.Value * lblItemC(0).Height
End Sub

Private Sub vsrChild_Scroll()
pikItems1.Top = vsrChild.Value * lblItemC(0).Height
End Sub

Private Sub vsrParent_Change()
pikItems.Top = vsrParent.Value * lblItemP(0).Height
End Sub
Sub ClearListChild()
Dim L As Long

On Error Resume Next

For L = 1 To lblItemC.ubound + 1
        Unload lblItemC(L)
Next L
End Sub
Sub GetChildrenOfSel()
    Dim L As Long, sClassRet As String
    
    ClearListChild
        
    Set ChildWins = New Collection
    
    lRetValue = EnumChildWindows(lblItemP(lSelIndex(1)).Tag, AddressOf EnumChildProc, 0)
    
        For L = 0 To ChildWins.Count - 1
            
                sClassRet = Space(255)
                lRetValue = GetClassName(ChildWins.Item(L + 1), sClassRet, 255)
                sClassRet = Left(sClassRet, lRetValue)
                            
                If Not (L = 0) Then
                    Load lblItemC(L)
                
                
                    With lblItemC(L)
                        .Left = 1
                        .Top = (lblItemC(L - 1).Height + lblItemC(L - 1).Top)
                        .Width = pikItems1.Width - 1
                        .Height = 270
                        .Tag = ChildWins.Item(L + 1)
                        .Caption = sClassRet
                        .Visible = True
                        .ForeColor = vbWhite
                    End With
                Else
                    With lblItemC(0)
                        .Left = 1
                        .Top = 0
                        .Width = pikItems1.Width - 1
                        .Height = 270
                        .Tag = ChildWins.Item(1)
                        .Caption = sClassRet
                        .ForeColor = vbWhite
                    End With
                End If
        Next L

pikItems1.Height = (lblItemC(0).Height * (lblItemC.Count)) + 45
vsrChild.Max = (pikFramed2.Height - pikItems1.Height) / (lblItemC(0).Height)
vsrChild.Visible = pikItems1.Height > pikFramed2.Height
pikFramed2.Height = (lblItemC(0).Height * 11) + 45
vsrChild.Height = pikFramed2.Height
pikItems1.Top = 0

    If ChildWins.Count = 0 Then
        lblItemC(0).Caption = "Window Has No Children"
        lblItemC(0).Tag = "NoGo"
        lblItemC(0).ForeColor = vbRed
    End If

End Sub

Private Sub vsrParent_Scroll()
pikItems.Top = vsrParent.Value * lblItemP(0).Height
End Sub
