VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Module1.ShowFind Me, FR_DOWN Or FR_SHOWHELP, "Find Text"
End Sub

Private Sub Command2_Click()
  Module1.ShowFind Me, FR_SHOWHELP, "Find Text", True, "Replace Text"
End Sub

Private Sub Form_Load()
   Caption = "Find/Replace dialogs"
   Command1.Caption = "Find dialog"
   Command2.Caption = "Replace dialog"
End Sub
