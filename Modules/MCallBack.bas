Attribute VB_Name = "MCallBack"
Option Explicit

'Public Function FolderBrowserDialogCallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal lParam As Long, ByVal lpData As Object) As Long
Public Function FolderBrowserDialogCallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If lpData = 0 Then Exit Function 'Is Nothing Then
    Dim fbd As FolderBrowserDialog: Set fbd = MObjPtr.PtrToObject(lpData)
    Dim icb As ICallBack: Set icb = fbd
    icb.CallBack hwnd, msg, lParam
    Set icb = Nothing
    MObjPtr.ZeroToObject fbd
End Function
