Attribute VB_Name = "Module1"
Option Explicit

' fabel358:
' https://www.vbforums.com/showthread.php?902963-Find-Replace-Dialog
' To Work, it works... but, too long!
' i don 't remember where I found this code; I found it among my material.


Type FINDREPLACE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    flags As Long
    lpstrFindWhat As Long
    lpstrReplaceWith As Long
    wFindWhatLen As Integer
    wReplaceWithLen As Integer
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

Private Declare Function FindText Lib "comdlg32" Alias "FindTextA" (pFindreplace As Long) As Long
Private Declare Function ReplaceText Lib "comdlg32" Alias "ReplaceTextA" (pFindreplace As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As Msg) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcessHeap& Lib "kernel32" ()
Private Declare Function HeapAlloc& Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long)
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Const GWL_WNDPROC = (-4)
Private Const HEAP_ZERO_MEMORY = &H8
Public Const FR_DIALOGTERM = &H40
Public Const FR_DOWN = &H1
Public Const FR_ENABLEHOOK = &H100
Public Const FR_ENABLETEMPLATE = &H200
Public Const FR_ENABLETEMPLATEHANDLE = &H2000
Public Const FR_FINDNEXT = &H8
Public Const FR_HIDEMATCHCASE = &H8000
Public Const FR_HIDEUPDOWN = &H4000
Public Const FR_HIDEWHOLEWORD = &H10000
Public Const FR_MATCHCASE = &H4
Public Const FR_NOMATCHCASE = &H800
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_REPLACE = &H10
Public Const FR_REPLACEALL = &H20
Public Const FR_SHOWHELP = &H80
Public Const FR_WHOLEWORD = &H2

Const FINDMSGSTRING = "commdlg_FindReplace"
Const HELPMSGSTRING = "commdlg_help"
Const BufLength = 256
Public hDialog As Long, OldProc As Long
Dim uFindMsg As Long, uHelpMsg As Long, lHeap As Long
Public RetFrs As FINDREPLACE, TMsg As Msg
Dim arrFind() As Byte, arrReplace() As Byte

'Private m_FRS As FINDREPLACE

Public Sub ShowFind(fOwner As Form, lFlags As Long, sFind As String, Optional bReplace As Boolean = False, Optional sReplace As String = "")
    If hDialog > 0 Then Exit Sub
    Dim FRS As FINDREPLACE
    Dim i As Integer
    arrFind = StrConv(sFind & Chr$(0), vbFromUnicode)
    arrReplace = StrConv(sReplace & Chr$(0), vbFromUnicode)
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
    CopyMemory ByVal lHeap, FRS, Len(FRS)
    uFindMsg = RegisterWindowMessage(FINDMSGSTRING)
    uHelpMsg = RegisterWindowMessage(HELPMSGSTRING)
    OldProc = SetWindowLong(fOwner.hwnd, GWL_WNDPROC, AddressOf WndProc)
    If bReplace Then
        hDialog = ReplaceText(ByVal lHeap)
    Else
        hDialog = FindText(ByVal lHeap)
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
        CopyMemory RetFrs, ByVal lParam, Len(RetFrs)
        If (RetFrs.flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
           SetWindowLong hOwner, GWL_WNDPROC, OldProc
           HeapFree GetProcessHeap(), 0, lHeap
           hDialog = 0: lHeap = 0: OldProc = 0
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
    Dim sTemp As String
    sTemp = "Here is your code for Find/Replace with parameters:" & vbCrLf & vbCrLf
    sTemp = sTemp & "Find string: " & PointerToString(fr.lpstrFindWhat) & vbCrLf
    sTemp = sTemp & "Replace string: " & PointerToString(fr.lpstrReplaceWith) & vbCrLf & vbCrLf
    sTemp = sTemp & "Current Flags: " & vbCrLf & vbCrLf
    sTemp = sTemp & "FR_FINDNEXT = " & CheckFlags(FR_FINDNEXT, fr.flags) & vbCrLf
    sTemp = sTemp & "FR_REPLACE = " & CheckFlags(FR_REPLACE, fr.flags) & vbCrLf
    sTemp = sTemp & "FR_REPLACEALL = " & CheckFlags(FR_REPLACEALL, fr.flags) & vbCrLf
    sTemp = sTemp & "FR_DOWN = " & CheckFlags(FR_DOWN, fr.flags) & vbCrLf
    sTemp = sTemp & "FR_MATCHCASE = " & CheckFlags(FR_MATCHCASE, fr.flags) & vbCrLf
    sTemp = sTemp & "FR_WHOLEWORD = " & CheckFlags(FR_WHOLEWORD, fr.flags)
    MsgBox sTemp, vbOKOnly + vbInformation, "Find/Replace parameters"
End Sub

Private Function PointerToString(p As Long) As String
    Dim s As String
    s = String(BufLength, Chr$(0))
    CopyPointer2String s, p
    PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Function CheckFlags(flag As Long, flags As Long) As Boolean
    CheckFlags = ((flags And flag) = flag)
End Function
