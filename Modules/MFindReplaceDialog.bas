Attribute VB_Name = "MFindReplaceDialog"
Option Explicit

'Hooks
'https://learn.microsoft.com/de-de/windows/win32/winmsg/hooks

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/ns-winuser-msg
'typedef struct tagMSG {
'  HWND   hwnd;
'  UINT   message;
'  WPARAM wParam;
'  LPARAM lParam;
'  DWORD  time;
'  POINT  pt;
'  DWORD  lPrivate;
'} MSG, *PMSG, *NPMSG, *LPMSG;

Public Type msg
    hwnd     As LongPtr
    message  As Long
    wParam   As LongPtr
    lParam   As LongPtr
    time     As Long
    ptX      As Long
    ptY      As Long
    'lPrivate As Long 'nur auf MAC
End Type

'typedef struct tagCWPRETSTRUCT {
'  LRESULT lResult;
'  LPARAM  lParam;
'  WPARAM  wParam;
'  UINT    message;
'  HWND    hwnd;
'} CWPRETSTRUCT, *PCWPRETSTRUCT, *NPCWPRETSTRUCT, *LPCWPRETSTRUCT;
'
'Private Type CWPRETSTRUCT
'    lResult As LongPtr
'    lParam  As LongPtr
'    wParam  As Long
'    message As Long
'    hwnd    As LongPtr
'End Type

Private Const WH_MSGFILTER       As Long = -1 ' Installiert eine Hookprozedur, die nachrichten überwacht, die als Ergebnis eines Eingabeereignisses in einem Dialogfeld, Meldungsfeld, Menü oder Einer Bildlaufleiste generiert werden. Weitere Informationen finden Sie in der Hookprozedur MessageProc .
Private Const WH_JOURNALRECORD   As Long = 0  ' Warnung 'Journaling Hooks-APIs werden ab Windows 11 nicht mehr unterstützt und werden in einer zukünftigen Version entfernt. Daher wird dringend empfohlen, stattdessen die SendInput TextInput-API aufzurufen. Installiert eine Hookprozedur, die Eingabenachrichten aufzeichnet, die an die Systemnachrichtenwarteschlange gesendet werden. Dieser Hook ist nützlich für die Aufzeichnung von Makros. Weitere Informationen finden Sie in der Hookprozedur JournalRecordProc .
Private Const WH_JOURNALPLAYBACK As Long = 1  ' Warnung 'Journaling Hooks-APIs werden ab Windows 11 nicht mehr unterstützt und werden in einer zukünftigen Version entfernt. Daher wird dringend empfohlen, stattdessen die SendInput TextInput-API aufzurufen. Installiert eine Hookprozedur, die zuvor von einer WH_JOURNALRECORD Hookprozedur aufgezeichnete Nachrichten veröffentlicht. Weitere Informationen finden Sie in der Hookprozedur JournalPlaybackProc .
Private Const WH_KEYBOARD        As Long = 2  ' Installiert eine Hookprozedur, die Tastatureingabemeldungen überwacht. Weitere Informationen finden Sie unter KeyboardProc-Hookprozedur .
Private Const WH_GETMESSAGE      As Long = 3  ' Installiert eine Hookprozedur, die Nachrichten überwacht, die an eine Nachrichtenwarteschlange gesendet werden. Weitere Informationen finden Sie in der Hookprozedur GetMsgProc .
Private Const WH_CALLWNDPROC     As Long = 4  ' Installiert eine Hookprozedur, die Nachrichten überwacht, bevor das System sie an die Zielfensterprozedur sendet. Weitere Informationen finden Sie in der Hookprozedur CallWndProc .
Private Const WH_CBT             As Long = 5  ' Installiert eine Hookprozedur, die Benachrichtigungen empfängt, die für eine CBT-Anwendung nützlich sind. Weitere Informationen finden Sie in der CBTProc-Hookprozedur .
Private Const WH_SYSMSGFILTER    As Long = 6  ' Installiert eine Hookprozedur, die nachrichten überwacht, die als Ergebnis eines Eingabeereignisses in einem Dialogfeld, Meldungsfeld, Menü oder Einer Bildlaufleiste generiert werden. Die Hookprozedur überwacht diese Meldungen für alle Anwendungen auf demselben Desktop wie der aufrufende Thread. Weitere Informationen finden Sie in der Hookprozedur SysMsgProc .
Private Const WH_MOUSE           As Long = 7  ' Installiert eine Hookprozedur, die Mausnachrichten überwacht. Weitere Informationen finden Sie unter MouseProc-Hookprozedur .
'8 ???
Private Const WH_DEBUG           As Long = 9  ' Installiert eine Hookprozedur, die zum Debuggen anderer Hookprozeduren nützlich ist. Weitere Informationen finden Sie in der DebugProc-Hookprozedur .
Private Const WH_SHELL           As Long = 10 ' Installiert eine Hookprozedur, die Benachrichtigungen empfängt, die für Shellanwendungen nützlich sind. Weitere Informationen finden Sie in der ShellProc-Hookprozedur .
Private Const WH_FOREGROUNDIDLE  As Long = 11 ' Installiert eine Hookprozedur, die aufgerufen wird, wenn sich der Vordergrundthread der Anwendung im Leerlauf befindet. Dieser Hook ist nützlich, um Aufgaben mit niedriger Priorität während der Leerlaufzeit auszuführen. Weitere Informationen finden Sie in der Hookprozedur ForegroundIdleProc .
Private Const WH_CALLWNDPROCRET  As Long = 12 ' Installiert eine Hookprozedur, die Nachrichten überwacht, nachdem sie von der Zielfensterprozedur verarbeitet wurden. Weitere Informationen finden Sie in der HookPROC-Rückruffunktions-Hookprozedur .
Private Const WH_KEYBOARD_LL     As Long = 13 ' Installiert eine Hookprozedur, die Tastatureingabeereignisse auf niedriger Ebene überwacht. Weitere Informationen finden Sie in der Hookprozedur LowLevelKeyboardProc .
Private Const WH_MOUSE_LL        As Long = 14 ' Installiert eine Hookprozedur, die Mauseingabeereignisse auf niedriger Ebene überwacht. Weitere Informationen finden Sie in der Hookprozedur LowLevelMouseProc .

    
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-setwindowshookexw
'HHOOK SetWindowsHookExW(  [in] int       idHook,  [in] HOOKPROC  lpfn,  [in] HINSTANCE hmod,  [in] DWORD     dwThreadId );
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-callnexthookex
'LRESULT CallNextHookEx( [in, optional] HHOOK hhk, [in] int nCode, [in] WPARAM wParam, [in] LPARAM lParam);
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long


'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-unhookwindowshookex
'BOOL UnhookWindowsHookEx(  [in] HHOOK hhk );
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nc-winuser-hookproc
'LRESULT Hookproc(      int code,  [in] WPARAM wParam,   [in] LPARAM lParam )


Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)

'https://learn.microsoft.com/de-de/windows/win32/api/processthreadsapi/nf-processthreadsapi-getcurrentthreadid
Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private m_hHook As LongPtr
Public mMSG     As msg

Function GetHInstance() As LongPtr
    'wird hier nicht gebraucht
#If VBA7 Then
    GetHInstance = Excel.Application.hInstancePtr
#Else
    GetHInstance = App.hInstance
#End If
End Function

Public Sub HookIt()
    'm_HHook = SetWindowsHookExW(WH_MSGFILTER, AddressOf MessageProc, GetHInstance, GetCurrentThreadId)
    'The hMod parameter must be set to NULL if the dwThreadId parameter specifies a thread created by the current process
    m_hHook = SetWindowsHookExW(WH_MSGFILTER, AddressOf MessageProc, 0&, GetCurrentThreadId)
End Sub

Public Sub UnHookIt()
    Dim hr As Long: hr = UnhookWindowsHookEx(m_hHook)
End Sub

'Private Function MessageProc(ByVal code As Long, ByVal wParam As LongPtr, ByRef lParam As CWPRETSTRUCT) As Long
'Private Function MessageProc(ByVal code As Long, ByVal wParam As LongPtr, ByRef lParam As MSG) As Long
Private Function MessageProc(ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    
    If lParam = 0 Then Exit Function
    RtlMoveMemory mMSG, ByVal lParam, LenB(mMSG)
    
    'Übergibt die Hookinformationen an die nächste Hookprozedur in der aktuellen Hookkette
    Dim lr As Long: lr = CallNextHookEx(m_hHook, code, wParam, lParam)
    If code < 0 Then
        MessageProc = 1 'irgendwas ungleich 0 ?
    Else
        'return
        Exit Function
    End If
End Function
