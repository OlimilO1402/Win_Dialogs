VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim PFN As String
    If MFileDlg.OpenFile_ShowDialog(Me.hWnd, App.Path, ".txt", "Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*", "Datei Öffnen", PFN) = vbCancel Then Exit Sub
    Debug.Print PFN
End Sub

Private Sub Command2_Click()
    Dim PFN As String
    If MFileDlg.OpenFile_ShowDialog2(Me.hWnd, App.Path, ".txt", "Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*", "Datei Öffnen", PFN) = vbCancel Then Exit Sub
    Debug.Print PFN
End Sub

Private Sub Command3_Click()
    Dim PFN As String
    If MFileDlg.OpenFile_ShowDialog3(Me.hWnd, App.Path, ".txt", "Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*", "Datei Öffnen", PFN) = vbCancel Then Exit Sub
    Debug.Print PFN
End Sub
