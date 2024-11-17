Attribute VB_Name = "MCallBack"
Option Explicit
Private Const WM_USER              As Long = &H400&
Private Const BFFM_SETSTATUSTEXTA  As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK        As Long = (WM_USER + 101)  '1125
Private Const BFFM_SETSELECTIONA   As Long = (WM_USER + 102)  '1126
Private Const BFFM_SETSELECTIONW   As Long = (WM_USER + 103)  '1127
Private Const BFFM_SETSTATUSTEXTW  As Long = (WM_USER + 104)  '1128

Private Const BFFM_INITIALIZED     As Long = 1
Private Const BFFM_SELCHANGED      As Long = 2
'Private Const BFFM_VALIDATEFAILEDA As Long = 3
'Private Const BFFM_VALIDATEFAILEDW As Long = 4
#If VBA7 Then
    Private Declare PtrSafe Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As LongPtr) As Long
    Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
    Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As LongPtr) As Long
    Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#End If

'Public Function FolderBrowserDialogCallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal lParam As Long, ByVal lpData As Object) As Long
Public Function FolderBrowserDialogCallBack(ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
'    'If lpData = 0 Then Exit Function 'Is Nothing Then
'    'Dim fbd As FolderBrowserDialog: Set fbd = MObjPtr.CObj(lpData)
'    'Dim icb As ICallBack: Set icb = fbd
'    'icb.CallBack hwnd, msg, lParam
'    'Set icb = Nothing
'    'MObjPtr.ZeroObj fbd
'    If Not lpData Is Nothing Then
'        If TypeOf lpData Is ICallBack Then
'            Call CCallBack(lpData).CallBack(hwnd, msg, lParam)
'        End If
'    End If
Try: On Error GoTo Catch


    'Dim rv As LongPtr
    Dim hr As LongPtr
    
    Select Case Msg
    Case BFFM_INITIALIZED
        'If (Len(mSelectedPath) > 0) Then
        hr = SendMessageW(hwnd, BFFM_SETSELECTIONW, 1&, ByVal lpData)
        'End If
    Case BFFM_SELCHANGED
        If (lParam <> 0&) Then
            Dim Buffer As String: Buffer = String$(1024, vbNullChar)
            hr = SHGetPathFromIDListW(lParam, ByVal StrPtr(Buffer))
            If hr = 1 Then
                hr = SendMessageW(hwnd, BFFM_ENABLEOK, 0, ByVal 1)
                hr = SendMessageW(hwnd, BFFM_SETSTATUSTEXTA, 0, StrPtr(Buffer))
            ElseIf hr = 0 Then
                hr = SendMessageW(hwnd, BFFM_ENABLEOK, 0, ByVal 0)
            End If
            'CoTaskMemFree VarPtr(lParam)
        End If
    End Select

Catch:


'Code by CallunWillock:
'
'  Private Function BrowseCallbackProc(ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal Pointer As LongPtr, ByVal Data As LongPtr) As LongPtr
'    On Error Resume Next
'
'    Dim Result      As Long
'    Dim Buffer      As String
'
'    Select Case Msg
'    Case BFFM_INITIALIZED
'        Call SendMessageW(hwnd, BFFM_SETSELECTION, 1&, Data)
'    Case BFFM_SELCHANGED
'        Buffer = Space(MAX_PATH)
'        Result = SHGetPathFromIDListW(Pointer, StrPtr(Buffer))
'        If Result = 1 Then
'          Call SendMessageW(hwnd, BFFM_SETSTATUSTEXTA, 0, StrPtr(Buffer))
'        End If
'    End Select
'    BrowseCallbackProc = 0
'  End Function



End Function

'Public Function CCallBack(ByVal obj As Object) As ICallBack
'    Set CCallBack = obj
'End Function
'
'  HWND unnamedParam1,
'  UINT unnamedParam2,
'  WPARAM unnamedParam3,
'  lParam unnamedParam4

Public Function FindReplaceCallBack(ByVal param1 As LongPtr, ByVal param2 As LongPtr, ByVal param3 As LongPtr, ByVal param4 As LongPtr)
    '
    
End Function
