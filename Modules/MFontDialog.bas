Attribute VB_Name = "MFontDialog"
Option Explicit
Public Const LF_FACESIZE    As Long = 32
Public Type LOGFONT 'W
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Integer 'String * LF_FACESIZE '(1 To LF_FACESIZE) As Byte
End Type

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

#If VBA7 Then
Private Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As Long
Private Declare PtrSafe Function SendDlgItemMessageW Lib "user32.dll" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As Long
Private Declare Function SendDlgItemMessageW Lib "user32.dll" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#End If

Public FontDialog As MyFontDialog
Dim defaultControlHwnd As LongPtr

'/// <summary>Gibt für Standarddialogfelder die Hookprozedur an, die überschrieben wird, um einem Standarddialogfeld bestimmte Funktionen hinzuzufügen.</summary>
'/// <returns>Ein Wert von 0, wenn die Meldung von der Prozedur für Standarddialogfelder verarbeitet wird. Ein Wert ungleich 0, wenn die Meldung von dieser Prozedur ignoriert wird.</returns>
'/// <param name="hWnd">Das Handle für das Dialogfeldfenster. </param>
'/// <param name="msg">Die empfangene Meldung. </param>
'/// <param name="wparam">Zusätzliche Informationen zur Meldung. </param>
'/// <param name="lparam">Zusätzliche Informationen zur Meldung. </param>
'[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
'    private  override IntPtr HookProc(IntPtr hWnd, int msg, IntPtr wparam, IntPtr lparam)
'    {
'        Switch (msg)
'        {
'        Case 273:
'        {
'            if ((int)wparam != 1026)
'            {
'                break
'            }
'            NativeMethods.LOGFONT lOGFONT = new NativeMethods.LOGFONT()
'            UnsafeNativeMethods.SendMessage(new HandleRef(null, hWnd), 1025, 0, lOGFONT)
'            UpdateFont (LOGFONT)
'            int num = (int)UnsafeNativeMethods.SendDlgItemMessage(new HandleRef(null, hWnd), 1139, 327, IntPtr.Zero, IntPtr.Zero)
'            if (num != -1)
'            {
'                UpdateColor((int)UnsafeNativeMethods.SendDlgItemMessage(new HandleRef(null, hWnd), 1139, 336, (IntPtr)num, IntPtr.Zero))
'            }
'            if (NativeWindow.WndProcShouldBeDebuggable)
'            {
'                OnApply (EventArgs.Empty)
'                break
'            }
'            Try
'            {
'                OnApply (EventArgs.Empty)
'            }
'            catch (Exception t)
'            {
'                Application.OnThreadException (t)
'            }
'            break
'        }
'        Case 272:
'            if (!showColor)
'            {
'                IntPtr dlgItem = UnsafeNativeMethods.GetDlgItem(new HandleRef(null, hWnd), 1139)
'                SafeNativeMethods.ShowWindow(new HandleRef(null, dlgItem), 0)
'                dlgItem = UnsafeNativeMethods.GetDlgItem(new HandleRef(null, hWnd), 1091)
'                SafeNativeMethods.ShowWindow(new HandleRef(null, dlgItem), 0)
'            }
'            break
'        }
'        return base.HookProc(hWnd, msg, wparam, lparam)
'    }
Public Function FncPtr(ByVal pAddr As LongPtr) As LongPtr
    FncPtr = pAddr
End Function

'private  override IntPtr HookProc(IntPtr hWnd, int msg, IntPtr wparam, IntPtr lparam)
Public Function FontDialog_HookProc(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case msg
    Case 273:
        If CLng(wParam) <> 1026 Then
            'break
            'Exit Function
        Else
            Dim aLogFont As LOGFONT
            SendMessageW hwnd, 1025, 0, aLogFont
            'sodale wie bekommt man jetzt den Logfont in den Dialog? OK, UpdateFont muss Friend sein
            FontDialog.UpdateFont aLogFont
            Dim hr As Long: hr = SendDlgItemMessageW(hwnd, 1139, 327, 0, 0)
            If hr <> -1 Then
                FontDialog.UpdateColor CLng(SendDlgItemMessageW(hwnd, 1139, 336, hr, 0))
            End If
        End If
    Case 272
        If Not FontDialog.ShowColor Then
            Dim dlgItem As LongPtr: dlgItem = GetDlgItem(hwnd, 1139)
            ShowWindow dlgItem, 0
            dlgItem = GetDlgItem(hwnd, 1091)
            ShowWindow dlgItem, 0
        End If
    End Select
    FontDialog_HookProc = CDBase_HookProc(hwnd, msg, wParam, lParam)
End Function

'    /// <summary>Definiert die Hookprozedur für Standarddialogfelder, die überschrieben wird, um einem Standarddialogfeld spezifische Funktionen hinzuzufügen.</summary>
'    /// <returns>Ein Wert von 0, wenn die Meldung von der Prozedur für Standarddialogfelder verarbeitet wird. Ein Wert ungleich 0, wenn die Meldung von dieser Prozedur ignoriert wird.</returns>
'    /// <param name="hWnd">Das Handle für das Dialogfeldfenster. </param>
'    /// <param name="msg">Die empfangene Meldung. </param>
'    /// <param name="wparam">Zusätzliche Informationen zur Meldung. </param>
'    /// <param name="lparam">Zusätzliche Informationen zur Meldung. </param>
'    [SecurityPermission(SecurityAction.InheritanceDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
'    [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
'    protected virtual IntPtr HookProc(IntPtr hWnd, int msg, IntPtr wparam, IntPtr lparam)
'    {
'        Switch (msg)
'        {
'        Case 272:
'            MoveToScreenCenter(hWnd);
'            defaultControlHwnd = wparam;
'            UnsafeNativeMethods.SetFocus(new HandleRef(null, wparam));
'            break;
'        Case 7:
'            UnsafeNativeMethods.PostMessage(new HandleRef(null, hWnd), 1105, 0, 0);
'            break;
'        Case 1105:
'            UnsafeNativeMethods.SetFocus(new HandleRef(this, defaultControlHwnd));
'            break;
'        }
'        return IntPtr.Zero;
'    }
'
'    internal static void MoveToScreenCenter(IntPtr hWnd)
'    {
'        NativeMethods.RECT rect = default(NativeMethods.RECT);
'        UnsafeNativeMethods.GetWindowRect(new HandleRef(null, hWnd), ref rect);
'        Rectangle workingArea = Screen.GetWorkingArea(Control.MousePosition);
'        int x = workingArea.X + (workingArea.Width - rect.right + rect.left) / 2;
'        int y = workingArea.Y + (workingArea.Height - rect.bottom + rect.top) / 3;
'        SafeNativeMethods.SetWindowPos(new HandleRef(null, hWnd), NativeMethods.NullHandleRef, x, y, 0, 0, 21);
'    }

Private Function CDBase_HookProc(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case msg
    Case 272
        MoveToScreenCenter hwnd
        defaultControlHwnd = wParam
        SetFocus wParam
    Case 7:
        PostMessageW hwnd, 1105, 0, 0
    Case 1105:
        SetFocus defaultControlHwnd
    End Select
    CDBase_HookProc = 0
End Function

Private Sub MoveToScreenCenter(ByVal hwnd As LongPtr)
    '
End Sub
