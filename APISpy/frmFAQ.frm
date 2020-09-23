VERSION 5.00
Begin VB.Form frmFAQ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FAQ"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmFAQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFAQs 
      Height          =   645
      ItemData        =   "frmFAQ.frx":08CA
      Left            =   0
      List            =   "frmFAQ.frx":08CC
      TabIndex        =   0
      Top             =   -15
      Width           =   5130
   End
   Begin VB.Image imgCool 
      Height          =   2940
      Left            =   0
      Picture         =   "frmFAQ.frx":08CE
      Stretch         =   -1  'True
      Top             =   645
      Width           =   1755
   End
End
Attribute VB_Name = "frmFAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

lstFAQs.AddItem "Why Does The Message Create Cause An Illegal Operation?"
lstFAQs.AddItem "What Is The Point Of This FAQ Window?"
End Sub
