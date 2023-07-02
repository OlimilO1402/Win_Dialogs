VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskDialogSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://www.vbarchiv.net/tipps/tipp_2224-vista-taskdialogindirect.html

'TASKDIALOG_COMMON_BUTTON_FLAGS
'Private Const TDCBF_OK_BUTTON     As Long = 1  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: OK.
'Private Const TDCBF_YES_BUTTON    As Long = 2  ' Das Dialogfeld "Aufgabe" enthält die Schaltfläche "Push": Ja.
'Private Const TDCBF_NO_BUTTON     As Long = 4  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Nein.
'Private Const TDCBF_CANCEL_BUTTON As Long = 8  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Abbrechen.
'                                               ' Diese Schaltfläche muss für das Dialogfeld angegeben werden,
'                                               ' um auf typische Abbrechenaktionen (Alt-F4 und Escape) zu reagieren.
'Private Const TDCBF_RETRY_BUTTON  As Long = 16 ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Wiederholen Sie den Vorgang.
'Private Const TDCBF_CLOSE_BUTTON  As Long = 32 ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Schließen.
Public Enum ETaskDialogButton
    tdButtonOK = &H1&
    tdButtonYes = &H2&
    tdButtonNo = &H4&
    tdButtonCancel = &H8&
    tdButtonRetry = &H10&
    tdButtonClose = &H20&
End Enum

'Private Const TD_WARNING_ICON         As Integer = -1 ' Im Aufgabendialogfeld wird ein Ausrufezeichen angezeigt.
'Private Const TD_ERROR_ICON           As Integer = -2 ' Ein Stoppzeichensymbol wird im Aufgabendialogfeld angezeigt.
'Private Const TD_INFORMATION_ICON     As Integer = -3 ' Ein Symbol, das aus einem Kleinbuchstaben i in einem Kreis besteht, wird im Aufgabendialogfeld angezeigt.
'Private Const TD_SHIELD_ICON          As Integer = -4 ' Ein Sicherheitsschildsymbol wird im Aufgabendialogfeld angezeigt.
'Public Const TD_SECURITY_ICON_BLUE    As Integer = -5 ' MAKEINTRESOURCEW(-5)
'Public Const TD_SECURITY_ICON_WARNING As Integer = -6
'Public Const TD_SECURITY_ICON_ERROR   As Integer = -7
'Public Const TD_SECURITY_ICON_OK      As Integer = -8
'Public Const TD_SECURITY_ICON_GREY    As Integer = -9
'Public Const TD_QUESTION_ICON         As Integer = &H7F01 '254 '32514

Public Enum ETaskDialogIcon
    tdIconWarning = -1
    tdIconError = -2
    tdIconInformation = -3
    tdIconShield = -4
    tdIconSecurityBlue = -5
    tdIconSecurityWarning = -6
    tdIconSecurityError = -7
    tdIconSecurityOK = -8
    tdIconSecurityGrey = -9
    tdIconQuestion = -33022
End Enum

'0   Fehler beim Funktionsaufruf. Weitere Informationen finden Sie im Verweis auf den Rückgabewert.
'gleich wie VbMsgBoxResult
Private Const IDOK     As Long = 1 ' Schaltfläche    'OK ' wurde ausgewählt.
Private Const IDCANCEL As Long = 2 ' Schaltfläche "Abbrechen" wurde ausgewählt, ALT-F4 wurde gedrückt, Escape wurde gedrückt, oder der Benutzer auf die Schaltfläche "Fenster schließen" geklickt.

Private Const IDRETRY  As Long = 4 ' Die Schaltfläche "Wiederholen" wurde ausgewählt.

Private Const IDYES    As Long = 6 ' Ja-Schaltfläche wurde ausgewählt.
Private Const IDNO     As Long = 7 ' Es wurde keine Schaltfläche ausgewählt.
Private Const IDCLOSE  As Long = 8 ' Es wurde die Schließen Schaltfläche ausgewählt.

Private Const S_OK          As Long = 0          ' Der Vorgang wurde erfolgreich abgeschlossen.
Private Const E_FAIL        As Long = &H80004005 ' Fehler beim Vorgang.
Private Const E_OUTOFMEMORY As Long = &H8007000E ' Es gibt nicht genügend Arbeitsspeicher, um den Vorgang abzuschließen.
Private Const E_INVALIDARG  As Long = &H80070057 ' Mindestens ein Argument ist ungültig.
Private Const E_ICONERROR1  As Long = &H80070715 ' The specified resource type can not be found in the image file.
Private Const E_ICONERROR2  As Long = &H80070716 ' The specified resource name can not be found in the image file.

'HRESULT TaskDialog(
'  [in]  HWND                           hwndOwner,
'  [in]  HINSTANCE                      hInstance,
'  [in]  PCWSTR                         pszWindowTitle,
'  [in]  PCWSTR                         pszMainInstruction,
'  [in]  PCWSTR                         pszContent,
'  [in]  TASKDIALOG_COMMON_BUTTON_FLAGS dwCommonButtons,
'  [in]  PCWSTR                         pszIcon,
'  [out] int                            *pnButton
');
#If VBA7 Then
    Private Declare PtrSafe Function TaskDialogSimple Lib "comctl32" Alias "TaskDialog" ( _
        ByVal hwndOwner As LongPtr, ByVal hInstance As LongPtr, _
        ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, _
        ByVal CommonButtons As Long, ByVal pszIcon As LongPtr, ByRef pnButton_out As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function TaskDialogSimple Lib "comctl32" Alias "TaskDialog" ( _
        ByVal hWndOwner As LongPtr, ByVal hInstance As LongPtr, _
        ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, _
        ByVal CommonButtons As Long, ByVal pszIcon As LongPtr, ByRef pnButton_out As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As LongPtr
#End If
Private m_Title   As String
Private m_Instr   As String
Private m_Content As String
Private m_Icon    As Variant
Private m_Buttons As Long

Friend Sub New_(Title As String, Instruction As String, Content As String, Optional ByVal Icon As ETaskDialogIcon, Optional ByVal Buttons As ETaskDialogButton)
    m_Title = Title: m_Instr = Instruction: m_Content = Content: m_Icon = Icon: m_Buttons = Buttons
End Sub

'Copy to MNew
'Public Function TaskDialog(Title As String, Instruction As String, Content As String, Optional ByVal Icon As ETaskDialogIcon, Optional ByVal Buttons As ETaskDialogButton) As TaskDialog
'    Set TaskDialog = New TaskDialog: TaskDialog.New_ Title, Instruction, Content, Icon, Buttons
'End Function
#If VBA Then
Public Function ShowDialog(Optional Owner As UserForm) As VbMsgBoxResult
    Dim hWndOwner As LongPtr: hWndOwner = GetActiveWindow
    Dim hInstance As LongPtr: hInstance = Excel.Application.hInstancePtr
End Function
#Else
Public Function ShowDialog(Optional Owner As Form) As VbMsgBoxResult
    Dim hWndOwner As LongPtr: hWndOwner = Owner.hwnd
    Dim hInstance As LongPtr: hInstance = App.hInstance
#End If
Try: On Error GoTo Finally
    Dim retButton As Long
    Dim hr As Long: hr = TaskDialogSimple(hWndOwner, 0, StrPtr(m_Title), StrPtr(m_Instr), StrPtr(m_Content), m_Buttons, MAKEINTRESOURCEW(m_Icon), retButton)
    ShowDialog = retButton
    If hr = S_OK Then Exit Function
Finally:
    Dim sErr As String
    Select Case hr
    Case E_FAIL:        sErr = "E_FAIL"        ' Fehler beim Vorgang.
    Case E_OUTOFMEMORY: sErr = "E_OUTOFMEMORY" ' Es gibt nicht genügend Arbeitsspeicher, um den Vorgang abzuschließen.
    Case E_INVALIDARG:  sErr = "E_INVALIDARG"  ' Mindestens ein Argument ist ungültig.
    Case E_ICONERROR1:  sErr = "The specified resource type can not be found in the image file."
    Case E_ICONERROR2:  sErr = "The specified resource name can not be found in the image file."
    End Select
    MsgBox "TaskDialogSE:ShowDialog " & hr & vbCrLf & sErr
End Function

Private Function MAKEINTRESOURCEW(VIcon) As LongPtr
    Dim vt As VbVarType: vt = VarType(VIcon)
    Dim pIcon As LongPtr
    Select Case vt
    Case vbString
        pIcon = StrPtr(m_Icon)
    'Case vbInteger: iIcon = CInt(m_Icon): pIcon = VarPtr(iIcon)
    Case vbInteger, vbLong
        Dim iIcon As Long: iIcon = VIcon
        If iIcon < 0 Then
            pIcon = &H10000 + iIcon
        Else
            pIcon = iIcon
        End If
    End Select
    MsgBox vt & " " & VIcon & " " & pIcon
    MAKEINTRESOURCEW = pIcon
End Function

Public Property Get Title() As String
    Title = m_Title
End Property
Public Property Let Title(ByVal Value As String)
    m_Title = Value
End Property

Public Property Get Instruction() As String
    Instruction = m_Instr
End Property
Public Property Let Instruction(ByVal Value As String)
    m_Instr = Value
End Property

Public Property Get Content() As String
    Content = m_Content
End Property
Public Property Let Content(ByVal Value As String)
    m_Content = Value
End Property

Public Property Get Icon()
    Icon = m_Icon
End Property
Public Property Let Icon(ByVal Value)
    m_Icon = Value
End Property

Public Property Get Buttons()
    Buttons = m_Buttons
End Property
Public Property Let Buttons(ByVal Value)
    m_Buttons = Value
End Property


