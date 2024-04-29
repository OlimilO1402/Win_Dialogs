VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "FMain"
   ScaleHeight     =   4860
   ScaleWidth      =   14685
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkHandleWindowPosChanged 
      Caption         =   "Handle WM_WINDOWPOSCHANGED"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton BtnSetToSmallerSize 
      Caption         =   "Set to 200x350"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton BtnSetToLargerSize 
      Caption         =   "Set to 800x600"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISubclassedWindow

Private Const MAXHEIGHT As Long = 500
Private Const MAXWIDTH  As Long = 600
Private Const MINHEIGHT As Long = 200
Private Const MINWIDTH  As Long = 300

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type WINDOWPOS
    hWnd  As LongPtr
    hWndInsertAfter As LongPtr
    x     As Long
    y     As Long
    cx    As Long
    cy    As Long
    Flags As Long
End Type

Private Const WM_SIZING       As Long = &H214&
Private Const WM_WINDOWPOSCHANGED As Long = &H47&
Private Const WMSZ_LEFT       As Long = 1
Private Const WMSZ_TOP        As Long = 3
Private Const WMSZ_TOPLEFT    As Long = 4
Private Const WMSZ_TOPRIGHT   As Long = 5
Private Const WMSZ_BOTTOMLEFT As Long = 7

Private Sub Form_Load()
    If Not SubclassWindow(Me.hWnd, Me, ESubclassID.escidFrmMain) Then
        Debug.Print "Subclassing failed!"
    End If
End Sub

Private Sub Form_Resize()
    Me.Caption = "Size: " & CStr(Me.ScaleX(Me.Width, Me.ScaleMode, ScaleModeConstants.vbPixels)) & "x" & CStr(Me.ScaleY(Me.Height, Me.ScaleMode, ScaleModeConstants.vbPixels))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not UnSubclassWindow(Me.hWnd, ESubclassID.escidFrmMain) Then
        Debug.Print "UnSubclassing failed!"
    End If
End Sub

Private Sub BtnSetToSmallerSize_Click()
    Me.Move Me.Left, Me.Top, Me.ScaleX(200, ScaleModeConstants.vbPixels, Me.ScaleMode), Me.ScaleY(350, ScaleModeConstants.vbPixels, Me.ScaleMode)
End Sub

Private Sub BtnSetToLargerSize_Click()
    Me.Move Me.Left, Me.Top, Me.ScaleX(800, ScaleModeConstants.vbPixels, Me.ScaleMode), Me.ScaleY(600, ScaleModeConstants.vbPixels, Me.ScaleMode)
End Sub

Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal scID As ESubclassID, ByRef bCallDefProc As Boolean) As Long
Try: On Error GoTo Catch
    Dim lRet As Long
    Select Case scID
    Case ESubclassID.escidFrmMain
        lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
        Debug.Print "FMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(scID)
    End Select
    Exit Function
Catch:
    Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(scID) & ": ", Err.Number, Err.Description
End Function

Private Function HandleMessage_Form(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByRef bCallDefProc As Boolean) As Long
    Dim lRet As Long

Try: On Error GoTo Catch
    Select Case uMsg
    Case WM_SIZING
        Dim tRect As RECT: RtlMoveMemory ByVal VarPtr(tRect), ByVal lParam, LenB(tRect)
        If tRect.Right - tRect.Left < MINWIDTH Then
            Select Case wParam
            Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT
                tRect.Left = tRect.Right - MINWIDTH
            Case Else
                tRect.Right = tRect.Left + MINWIDTH
            End Select
        ElseIf tRect.Right - tRect.Left > MAXWIDTH Then
            Select Case wParam
            Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT
                tRect.Left = tRect.Right - MAXWIDTH
            Case Else
                tRect.Right = tRect.Left + MAXWIDTH
            End Select
        End If
        If tRect.Bottom - tRect.Top < MINHEIGHT Then
            Select Case wParam
            Case WMSZ_TOPLEFT, WMSZ_TOP, WMSZ_TOPRIGHT
                tRect.Top = tRect.Bottom - MINHEIGHT
            Case Else
                tRect.Bottom = tRect.Top + MINHEIGHT
            End Select
        ElseIf tRect.Bottom - tRect.Top > MAXHEIGHT Then
            Select Case wParam
            Case WMSZ_TOPLEFT, WMSZ_TOP, WMSZ_TOPRIGHT
                tRect.Top = tRect.Bottom - MAXHEIGHT
            Case Else
                tRect.Bottom = tRect.Top + MAXHEIGHT
            End Select
        End If
        RtlMoveMemory ByVal lParam, ByVal VarPtr(tRect), LenB(tRect)
        
    Case WM_WINDOWPOSCHANGED
        If chkHandleWindowPosChanged.Value = vbChecked Then
            Dim tWindowPos As WINDOWPOS: RtlMoveMemory ByVal VarPtr(tWindowPos), ByVal lParam, LenB(tWindowPos)
            If tWindowPos.cx < MINWIDTH Then
                On Error Resume Next
                Me.Width = ScaleX(MINWIDTH, ScaleModeConstants.vbPixels, Me.ScaleMode)
            ElseIf tWindowPos.cx > MAXWIDTH Then
                On Error Resume Next
                Me.Width = ScaleX(MAXWIDTH, ScaleModeConstants.vbPixels, Me.ScaleMode)
            End If
            If tWindowPos.cy < MINHEIGHT Then
                On Error Resume Next
                Me.Height = ScaleY(MINHEIGHT, ScaleModeConstants.vbPixels, Me.ScaleMode)
            ElseIf tWindowPos.cy > MAXHEIGHT Then
                On Error Resume Next
                Me.Height = ScaleY(MAXHEIGHT, ScaleModeConstants.vbPixels, Me.ScaleMode)
            End If
        End If
    End Select
  
    Exit Function
Catch:
    Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
End Function
