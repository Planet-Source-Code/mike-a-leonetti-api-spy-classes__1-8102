VERSION 5.00
Begin VB.Form frmHelpLockMode 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API Spy Help - Lock Mode"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmHelpLockMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgMain 
      Height          =   2115
      Left            =   60
      Picture         =   "frmHelpLockMode.frx":08CA
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelpLockMode.frx":13B1
      ForeColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   3165
   End
End
Attribute VB_Name = "frmHelpLockMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call Frm_OnTop(frmHelpLockMode, True)
End Sub

