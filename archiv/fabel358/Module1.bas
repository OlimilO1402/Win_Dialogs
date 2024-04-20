Attribute VB_Name = "Module1"
Option Explicit

Type FINDREPLACE
    lStructSize      As Long
    hwndOwner        As Long
    hInstance        As Long
    flags            As Long
    lpstrFindWhat    As Long
    lpstrReplaceWith As Long
    wFindWhatLen     As Integer
    wReplaceWithLen  As Integer
    lCustData        As Long
    lpfnHook         As Long
    lpTemplateName   As Long 'String
End Type

Type Msg
    hwnd    As Long
    message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    ptX     As Long
    ptY     As Long
End Type

Private Declare Function FindTextW Lib "comdlg32" (pFindreplace As Long) As Long
Private Declare Function ReplaceTextW Lib "comdlg32" (pFindreplace As Long) As Long

Private Declare Function RegisterWindowMessageW Lib "user32" (ByVal lpString As Long) As Long
Private Declare Function DispatchMessageW Lib "user32" (lpMsg As Msg) As Long
Private Declare Function GetMessageW Lib "user32" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function IsDialogMessageW Lib "user32" (ByVal hDlg As Long, lpMsg As Msg) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal cbCopy As Long)
Private Declare Function lstrcpyW Lib "kernel32" (ByVal NewString As Long, ByVal OldString As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const HEAP_ZERO_MEMORY  As Long = &H8

Public Const FR_DOWN                 As Long = &H1
Public Const FR_WHOLEWORD            As Long = &H2
Public Const FR_MATCHCASE            As Long = &H4
Public Const FR_FINDNEXT             As Long = &H8
Public Const FR_REPLACE              As Long = &H10
Public Const FR_REPLACEALL           As Long = &H20
Public Const FR_DIALOGTERM           As Long = &H40
Public Const FR_SHOWHELP             As Long = &H80

Public Const FR_ENABLEHOOK           As Long = &H100
Public Const FR_ENABLETEMPLATE       As Long = &H200
Public Const FR_NOUPDOWN             As Long = &H400
Public Const FR_NOMATCHCASE          As Long = &H800

Public Const FR_NOWHOLEWORD          As Long = &H1000
Public Const FR_ENABLETEMPLATEHANDLE As Long = &H2000
Public Const FR_HIDEUPDOWN           As Long = &H4000
Public Const FR_HIDEMATCHCASE        As Long = &H8000
Public Const FR_HIDEWHOLEWORD        As Long = &H10000

Const FINDMSGSTRING As String = "commdlg_FindReplace"
Const HELPMSGSTRING As String = "commdlg_help"
Const BufLength As Long = 256

Public hDialog As Long
Public OldProc As Long

Dim uFindMsg As Long
Dim uHelpMsg As Long
Dim lHeap As Long

Public RetFrs As FINDREPLACE
Public TMsg As Msg

Dim arrFind() As Byte
Dim arrReplace() As Byte

Public Sub ShowFind(fOwner As Form, lFlags As Long, sFind As String, Optional bReplace As Boolean = False, Optional sReplace As String = "")
    If hDialog > 0 Then Exit Sub
    Dim FRS As FINDREPLACE
    Dim i As Integer
    arrFind = StrConv(sFind & Chr$(0), vbUnicode)
    Debug.Print arrFind
    arrReplace = StrConv(sReplace & Chr$(0), vbUnicode)
    Debug.Print arrReplace
    With FRS
        .lStructSize = LenB(FRS) '&H20     '
        .lpstrFindWhat = VarPtr(arrFind(0))
        .wFindWhatLen = BufLength
        .lpstrReplaceWith = VarPtr(arrReplace(0))
        .wReplaceWithLen = BufLength
        .hwndOwner = fOwner.hwnd
        .flags = lFlags
        .hInstance = App.hInstance
    End With
    lHeap = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, FRS.lStructSize)
    RtlMoveMemory ByVal lHeap, FRS, Len(FRS)
    uFindMsg = RegisterWindowMessageW(StrPtr(FINDMSGSTRING))
    uHelpMsg = RegisterWindowMessageW(StrPtr(HELPMSGSTRING))
    OldProc = SetWindowLongW(fOwner.hwnd, GWL_WNDPROC, AddressOf WndProc)
    If bReplace Then
        hDialog = ReplaceTextW(ByVal lHeap)
    Else
        hDialog = FindTextW(ByVal lHeap)
    End If
    MessageLoop
End Sub

Private Sub MessageLoop()
    Do While GetMessage(TMsg, 0&, 0&, 0&) And hDialog > 0
        If IsDialogMessage(hDialog, TMsg) = False Then
            TranslateMessage TMsg
            DispatchMessage TMsg
        End If
    Loop
End Sub

Public Function WndProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg
    Case uFindMsg
        RtlMoveMemory RetFrs, ByVal lParam, LenB(RetFrs)
        If (RetFrs.flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
            SetWindowLongW hOwner, GWL_WNDPROC, OldProc
            HeapFree GetProcessHeap(), 0, lHeap
            hDialog = 0
            lHeap = 0
            OldProc = 0
        Else
            DoFindReplace RetFrs
        End If
    Case uHelpMsg
        MsgBox "Here is your code to call your help file", vbInformation + vbOKOnly, "Heeeelp!!!!"
    Case Else
        WndProc = CallWindowProc(OldProc, hOwner, wMsg, wParam, lParam)
    End Select
End Function

Private Sub DoFindReplace(fr As FINDREPLACE)
    Dim s As String
    s = "Here is your code for Find/Replace with parameters:" & vbCrLf & vbCrLf
    s = s & "Find string: " & PointerToString(fr.lpstrFindWhat) & vbCrLf
    s = s & "Replace string: " & PointerToString(fr.lpstrReplaceWith) & vbCrLf & vbCrLf
    s = s & "Current Flags: " & vbCrLf & vbCrLf
    s = s & "FR_FINDNEXT = " & CheckFlags(FR_FINDNEXT, fr.flags) & vbCrLf
    s = s & "FR_REPLACE = " & CheckFlags(FR_REPLACE, fr.flags) & vbCrLf
    s = s & "FR_REPLACEALL = " & CheckFlags(FR_REPLACEALL, fr.flags) & vbCrLf
    s = s & "FR_DOWN = " & CheckFlags(FR_DOWN, fr.flags) & vbCrLf
    s = s & "FR_MATCHCASE = " & CheckFlags(FR_MATCHCASE, fr.flags) & vbCrLf
    s = s & "FR_WHOLEWORD = " & CheckFlags(FR_WHOLEWORD, fr.flags)
    MsgBox s, vbOKOnly + vbInformation, "Find/Replace parameters"
End Sub

Private Function PointerToString(p As Long) As String
    Dim s As String: s = String(BufLength, Chr$(0))
    lstrcpyW StrPtr(s), p
    PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Function CheckFlags(flag As Long, flags As Long) As Boolean
   CheckFlags = ((flags And flag) = flag)
End Function
