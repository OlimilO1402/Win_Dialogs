Attribute VB_Name = "MSubClassing"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If
Public Enum ESubclassID
    escidFrmMain = 1
    'escidFrmMainCmdOk
    '...
End Enum

'https://learn.microsoft.com/de-de/windows/win32/api/commctrl/

'https://learn.microsoft.com/de-de/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/commctrl/nf-commctrl-defsubclassproc
'LRESULT DefSubclassProc( [in] HWND hWnd, [in] UINT uMsg, [in] WPARAM wParam, [in] LPARAM lParam );
Public Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/commctrl/nf-commctrl-removewindowsubclass
'BOOL RemoveWindowSubclass( [in] HWND hWnd, [in] SUBCLASSPROC pfnSubclass, [in] UINT_PTR uIdSubclass );
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/commctrl/nc-commctrl-subclassproc
'SUBCLASSPROC Subclassproc; LRESULT Subclassproc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData ) {...}

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal sz As Long)

Public Function SubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
Try: On Error GoTo Catch
    Dim bCallDefProc As Boolean: bCallDefProc = True
    If dwRefData Then
        Dim SCWnd As ISubclassedWindow: Set SCWnd = GetObjectFromPointer(dwRefData)
        If Not (SCWnd Is Nothing) Then
            SubclassProc = SCWnd.HandleMessage(hWnd, uMsg, wParam, lParam, uIdSubclass, bCallDefProc)
        End If
    End If
    On Error Resume Next
    If bCallDefProc Then
        Dim lr As LongPtr: lr = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
    Exit Function
Catch:
    Debug.Print "Error in SubclassProc: ", Err.Number, Err.Description
End Function

Public Function SubclassWindow(ByVal hWnd As LongPtr, SCWnd As ISubclassedWindow, ByVal scID As ESubclassID) As Boolean
Try: On Error GoTo Catch
    SubclassWindow = SetWindowSubclass(hWnd, AddressOf MSubClassing.SubclassProc, scID, ObjPtr(SCWnd)) <> 0
    Exit Function
Catch:
    Debug.Print "Error in SubclassWindow: ", Err.Number, Err.Description
End Function

Public Function UnSubclassWindow(ByVal hWnd As LongPtr, ByVal scID As ESubclassID) As Boolean
Try: On Error GoTo Catch
    UnSubclassWindow = RemoveWindowSubclass(hWnd, AddressOf MSubClassing.SubclassProc, scID) <> 0
    Exit Function
Catch:
    Debug.Print "Error in Function UnSubclassWindow: ", Err.Number, Err.Description
End Function

' returns the object <pObj> points to
Private Function GetObjectFromPointer(ByVal pObj As LongPtr) As Object
    Dim Obj As Object: RtlMoveMemory ByVal VarPtr(Obj), ByVal VarPtr(pObj), LenB(pObj)
    Set GetObjectFromPointer = Obj
    RtlZeroMemory ByVal VarPtr(Obj), LenB(pObj)
End Function
