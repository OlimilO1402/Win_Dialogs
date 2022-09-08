Attribute VB_Name = "MCallBack"
Option Explicit

'Public Function FolderBrowserDialogCallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal lParam As Long, ByVal lpData As Object) As Long
Public Function FolderBrowserDialogCallBack(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal lParam As LongPtr, ByVal lpData As Object) As Long
    'If lpData = 0 Then Exit Function 'Is Nothing Then
    'Dim fbd As FolderBrowserDialog: Set fbd = MObjPtr.CObj(lpData)
    'Dim icb As ICallBack: Set icb = fbd
    'icb.CallBack hwnd, msg, lParam
    'Set icb = Nothing
    'MObjPtr.ZeroObj fbd
    If Not lpData Is Nothing Then
        If TypeOf lpData Is ICallBack Then
            Call CCallBack(lpData).CallBack(hwnd, msg, lParam)
        End If
    End If
End Function

Public Function CCallBack(ByVal obj As Object) As ICallBack
    Set CCallBack = obj
End Function

