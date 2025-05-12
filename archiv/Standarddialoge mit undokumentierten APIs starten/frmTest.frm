VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Test"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "ShowFormatDriveDlg"
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   2100
      Width           =   1635
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ShowRunDlg"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   1620
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ShowPicIconDlg"
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ShowShutDownDlg"
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowOpenDlg"
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1635
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

Private Sub Command1_Click()
    MsgBox ShowOpenDlg(Me, "C:\", "Alle Dateien|*.*", , "Test: Bitte wählen!")
End Sub

Private Sub Command2_Click()
    ShowShutDownDlg
End Sub

Private Sub Command3_Click()
    Dim File As String: File = "moricons.dll"
    Dim Nr As Long

    ShowPicIconDlg Me, File, Nr
    MsgBox "Datei=" & File & vbCrLf & "Index=" & Nr
End Sub

Private Sub Command4_Click()
    ShowRunDlg Me, True
End Sub

Private Sub Command5_Click()
    ShowFormatDriveDlg Me
End Sub
