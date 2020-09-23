VERSION 5.00
Begin VB.Form frmInputBox 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   1920
      Width           =   5355
   End
   Begin VB.Label lblOk 
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      Left            =   5430
      TabIndex        =   2
      Top             =   1950
      Width           =   285
   End
   Begin VB.Image imgMain 
      Height          =   1965
      Left            =   0
      Picture         =   "frmInputBox.frx":0000
      Top             =   0
      Width           =   1920
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1905
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   3870
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = &HFFFFFF
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = &HFFFFFF
End Sub

Private Sub lblOk_Click()
sRetValue = txtMain.Text

frmInputBox.Hide
Unload frmInputBox
End Sub

Private Sub lblOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = &HC0C0C0
End Sub

Private Sub lblText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = &HFFFFFF
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)

If bNoLetters Then
    If Not (IsNumeric(Chr(KeyAscii))) And (Not (KeyAscii = 8)) Then
        KeyAscii = 0
    End If
End If

End Sub

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = &HFFFFFF
End Sub
