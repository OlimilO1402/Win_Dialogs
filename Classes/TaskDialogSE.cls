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
Private Const TDCBF_OK_BUTTON     As Long = 1  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: OK.
Private Const TDCBF_YES_BUTTON    As Long = 2  ' Das Dialogfeld "Aufgabe" enthält die Schaltfläche "Push": Ja.
Private Const TDCBF_NO_BUTTON     As Long = 4  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Nein.
Private Const TDCBF_CANCEL_BUTTON As Long = 8  ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Abbrechen.
                                               ' Diese Schaltfläche muss für das Dialogfeld angegeben werden,
                                               ' um auf typische Abbrechenaktionen (Alt-F4 und Escape) zu reagieren.
Private Const TDCBF_RETRY_BUTTON  As Long = 16 ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Wiederholen Sie den Vorgang.
Private Const TDCBF_CLOSE_BUTTON  As Long = 32 ' Das Dialogfeld "Aufgabe" enthält die Pushschaltfläche: Schließen.


'Public Const TD_WARNING_ICON          As Integer = -1       ' MAKEINTRESOURCEW(-1)
'Public Const TD_ERROR_ICON            As Integer = -2         ' MAKEINTRESOURCEW(-2)
'Public Const TD_INFORMATION_ICON      As Integer = -3   ' MAKEINTRESOURCEW(-3)
'Public Const TD_SHIELD_ICON           As Integer = -4        ' MAKEINTRESOURCEW(-4)
'Public Const TD_SECURITY_ICON_BLUE    As Integer = -5
'Public Const TD_SECURITY_ICON_WARNING As Integer = -6
'Public Const TD_SECURITY_ICON_ERROR   As Integer = -7
'Public Const TD_SECURITY_ICON_OK      As Integer = -8
'Public Const TD_SECURITY_ICON_GREY    As Integer = -9
'Public Const TD_QUESTION_ICON         As Integer = 32514

Private Const TD_WARNING_ICON     As Integer = -1 ' Im Aufgabendialogfeld wird ein Ausrufezeichen angezeigt.
Private Const TD_ERROR_ICON       As Integer = -2 ' Ein Stoppzeichensymbol wird im Aufgabendialogfeld angezeigt.
Private Const TD_INFORMATION_ICON As Integer = -3 ' Ein Symbol, das aus einem Kleinbuchstaben i in einem Kreis besteht, wird im Aufgabendialogfeld angezeigt.
Private Const TD_SHIELD_ICON      As Integer = -4 ' Ein Sicherheitsschildsymbol wird im Aufgabendialogfeld angezeigt.

'0   Fehler beim Funktionsaufruf. Weitere Informationen finden Sie im Verweis auf den Rückgabewert.
'gleich wie VbMsgBoxResult
Private Const IDOK     As Long = 1 ' Schaltfläche    'OK ' wurde ausgewählt.
Private Const IDCANCEL As Long = 2 ' Schaltfläche "Abbrechen" wurde ausgewählt, ALT-F4 wurde gedrückt, Escape wurde gedrückt, oder der Benutzer auf die Schaltfläche "Fenster schließen" geklickt.

Private Const IDRETRY  As Long = 4 ' Die Schaltfläche "Wiederholen" wurde ausgewählt.

Private Const IDYES    As Long = 6 ' Ja-Schaltfläche wurde ausgewählt.
Private Const IDNO     As Long = 7 ' Es wurde keine Schaltfläche ausgewählt.
Private Const IDCLOSE  As Long = 8 ' Es wurde die Schließen Schaltfläche ausgewählt.

Private Const S_OK          As Long = 0 ' Der Vorgang wurde erfolgreich abgeschlossen.
Private Const E_OUTOFMEMORY As Long = 0 ' Es gibt nicht genügend Arbeitsspeicher, um den Vorgang abzuschließen.
Private Const E_INVALIDARG  As Long = 0 ' Mindestens ein Argument ist ungültig.
Private Const E_FAIL        As Long = 0 ' Fehler beim Vorgang.

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
    Private Declare PtrSafe Function TaskDialogFun Lib "comctl32" Alias "TaskDialog" ( _
        ByVal hwndOwner As LongPtr, ByVal hInstance As LongPtr, _
        ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, _
        ByVal CommonButtons As Long, ByVal pszIcon As LongPtr, ByRef pnButton_out As Long) As Long
#Else
    Private Declare Function TaskDialogFun Lib "comctl32" Alias "TaskDialog" ( _
        ByVal hwndOwner As LongPtr, ByVal hInstance As LongPtr, _
        ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, _
        ByVal CommonButtons As Long, ByVal pszIcon As LongPtr, ByRef pnButton_out As Long) As Long
#End If
Private m_Title   As String
Private m_Instr   As String
Private m_Content As String
'Private m_sIcon   As String
'Private m_iIcon   As Integer
Private m_Icon    As Variant
Private m_Buttons As Long

Friend Sub New_(Title As String, Instruction As String, Content As String, Optional Icon, Optional Buttons As VbMsgBoxResult)
    m_Title = Title: m_Instr = Instruction: m_Content = Content: m_Icon = Icon: m_Buttons = Buttons
End Sub

Public Function TaskDialog(ByVal Title As String, ByVal Instruction As String, ByVal Content As String, Optional ByVal Icon, Optional ByVal Buttons As VbMsgBoxResult) As TaskDialog
    Set TaskDialog = New TaskDialog: TaskDialog.New_ Title, Instruction, Content, Icon, Buttons
End Function

Public Function ShowDialog(Optional Owner As UserForm) As VbMsgBoxResult
    Dim retButton As Long
    Dim vt As VbVarType: vt = VarType(m_Icon)
    Dim iIcon As Integer, pIcon As LongPtr
    Select Case vt
    Case vbString:  pIcon = StrPtr(m_Icon)
    Case vbInteger: iIcon = CInt(m_Icon): pIcon = VarPtr(iIcon)
    End Select
    Dim hr As Long: hr = TaskDialogFun(0, 0, StrPtr(m_Title), StrPtr(m_Instr), StrPtr(m_Content), m_Buttons, pIcon, retButton)
    ShowDialog = retButton
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


