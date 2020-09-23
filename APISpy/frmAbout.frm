VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   1860
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5190
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStupid 
      Interval        =   25
      Left            =   3255
      Top             =   840
   End
   Begin VB.PictureBox pikLines 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   1575
      TabIndex        =   5
      ToolTipText     =   "Simple example of VB's ""LINE"" Function..."
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email...."
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label lblCompile 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled..."
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1650
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1650
      TabIndex        =   2
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lPosX As Long, bRorL As Boolean, sNewColor As String, sNewColor2 As String
Dim lTime1 As Long, lTime2 As Long, lDuration As Long
Sub NewColor()

Do

    sNewColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    sNewColor2 = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

Loop Until sNewColor <> sNewColor2

End Sub

Private Sub cmdOK_Click()
  Unload frmAbout
End Sub

Private Sub Form_Load()
    lDuration = 25
    Call Frm_OnTop(frmAbout, True)
    frmAbout.Caption = "About " & VERSION_NAME & "'s API Spy"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblCompile.Caption = "Compiled in Visual Basic 6"
    lblDescription.Caption = "Enjoy this program... screw windows up!"
    lblEmail.Caption = "Email:  InFeStEd@optonline.net"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HC000&

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload frmAbout

End Sub
Private Sub lblCompile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HC000&

End Sub

Private Sub lblDescription_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HC000&

End Sub

Private Sub lblEmail_Click()
Call NET_JumpToWebsite("mailto:InFeStEd@optonline.net")
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HFF00&

End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HC000&

End Sub

Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEmail.ForeColor = &HC000&


End Sub

Private Sub tmrStupid_Timer()

        If bRorL = True Then
            If lPosX >= pikLines.ScaleWidth Then
                bRorL = False
                Call NewColor
            End If
        Else
            If lPosX <= 0 Then
                bRorL = True
                Call NewColor
            End If
        End If

        If bRorL = True Then
            Let lPosX = lPosX + 200
        Else
            Let lPosX = lPosX - 200
        End If

        Line (pikLines.ScaleWidth / 2, pikLines.ScaleHeight)-(lPosX, 0), sNewColor
        Line (pikLines.ScaleWidth / 2, 0)-(lPosX, pikLines.ScaleHeight), sNewColor
        

End Sub
