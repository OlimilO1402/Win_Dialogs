Attribute VB_Name = "ModCallBack"
Option Explicit

Public Function FolderBrowserDialogCallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal lParam As Long, ByVal lpData As Object) As Long
    If Not lpData Is Nothing Then
        If TypeOf lpData Is ICallBack Then
            Call CCallBack(lpData).CallBack(hwnd, msg, lParam)
        End If
    End If
End Function

Public Function CCallBack(ByVal obj As Object) As ICallBack
    Set CCallBack = obj
End Function
