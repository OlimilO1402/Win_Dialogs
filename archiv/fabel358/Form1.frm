VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   8055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   480
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
    Command1.Caption = "Find"
    Command2.Caption = "Replace"
End Sub
