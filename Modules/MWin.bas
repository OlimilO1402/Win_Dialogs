Attribute VB_Name = "MWin"
Option Explicit

'typedef struct tagHELPINFO {
'  UINT      cbSize;
'  int       iContextType;
'  int       iCtrlId;
'  HANDLE    hItemHandle;
'  DWORD_PTR dwContextId;
'  POINT     MousePos;
'} HELPINFO, *LPHELPINFO;

Private Type HelpInfo
    cbSize       As Long
    iContextType As Long
    iCtrlId      As Long
    hItemHandle  As Long 'Ptr
    dwContextId  As Long
    MousePosX    As Long
    MousePosY    As Long
End Type

Private m_HelpInfo As HelpInfo

Public LastMsgBoxResult As String

Public Function MessageBoxCallBack(lpHelpInfo As HelpInfo) As Long
    m_HelpInfo = lpHelpInfo
End Function

Public Function HelpInfo_ToStr() As String
    Dim s As String
    With m_HelpInfo
        s = "HelpInfo{" & vbCrLf
        s = s & "    cbSize      : " & .cbSize & vbCrLf
        s = s & "    iContextType: " & .iContextType & vbCrLf
        s = s & "    iCtrlId     : " & .iCtrlId & vbCrLf
        s = s & "    hItemHandle : " & .hItemHandle & vbCrLf
        s = s & "    dwContextId : " & .dwContextId & vbCrLf
        s = s & "    MousePosX   : " & .MousePosX & vbCrLf
        s = s & "    MousePosY   : " & .MousePosY & vbCrLf
        s = s & "}"
    End With
    HelpInfo_ToStr = s
End Function
Public Function MsgBox(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As Variant, Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult
'    Dim mb As MessageBox: Set mb = New MessageBox
'    With mb
'        .MsgBoxFncType = vbNormal
'        .Prompt = Prompt
'        .Style = Buttons
'        If Not IsMissing(Title) Then .Title = Title
'        MsgBox = .Show
'        LastResult = .Result_ToStr
'    End With
    
    'oder so:
    Dim mb As New MessageBox
    MsgBox = mb.Show(Prompt, Buttons, Title, HelpFile, Context)
    LastMsgBoxResult = mb.Result_ToStr
End Function

