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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SpecialFolder_Personal As Long = &H5&   ' = CSIDL_PERSONAL

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pIdl As Long) As Long

'Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByRef pv As Any)


'VB 5/6-Tipp 0061: Windows-Verzeichnisse erfassen: 'http://www.activevb.de/tipps/vb6tipps/tipp0061.html

Private Sub Command1_Click()
    MsgBox GetSpecialFolder(SpecialFolder_Personal)
End Sub

Function GetSpecialFolder(ByVal spf As Long) As String
    Dim pIdl As Long
    If SHGetSpecialFolderLocation(Me.hWnd, SpecialFolder_Personal, pIdl) <> 0 Then Exit Function
    Dim m_Buffer As String: m_Buffer = String$(1024, vbNullChar)
    'If SHGetPathFromIDListA(pIdl, m_Buffer) = 0 Then Exit Sub
    If SHGetPathFromIDListW(pIdl, StrPtr(m_Buffer)) = 0 Then Exit Function
    CoTaskMemFree pIdl
    Dim l As Long: l = InStr(m_Buffer, vbNullChar) - 1
    If l <= 0 Then Exit Function
    GetSpecialFolder = Left$(m_Buffer, l)
End Function
