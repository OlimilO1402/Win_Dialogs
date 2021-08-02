VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const IDOK       As Long = 1  '    The selected button was OK
Private Const IDCANCEL   As Long = 2  '    The selected button was Cancel
Private Const IDABORT    As Long = 3  '    The selected button was Abort
Private Const IDRETRY    As Long = 4  '    The selected button was Retry
Private Const IDIGNORE   As Long = 5  '    The selected button was Ignore
Private Const IDYES      As Long = 6  '    The selected button was Yes
Private Const IDNO       As Long = 7  '    The selected button was No
Private Const IDTRYAGAIN As Long = 10 '    The selected button was Try-Again
Private Const IDCONTINUE As Long = 11 '    The selected button was Continue

Public Enum MsgBoxResult
    vbOK = 1         'OK
    vbCancel = 2     'Abbrechen
    vbAbort = 3      'Abbrechen
    vbRetry = 4      'Wiederholen
    vbIgnore = 5     'Ignorieren
    vbYes = 6        'Ja
    vbNo = 7         'Nein
    vbTryAgain = 10  'Wiederholen
    vbContinue = 11  'Weiter
End Enum

Private Const MB_USERICON          As Long = &H80&

'Table 1
'To indicate the buttons displayed in the message box, specify one of the following values.
Private Const MB_OK                As Long = &H0&    '    The message box contains one push button : OK. This is the default.
Private Const MB_OKCANCEL          As Long = &H1&    '    The message box contains two push buttons: OK and Cancel.
Private Const MB_ABORTRETRYIGNORE  As Long = &H2&    '    The message box contains three push buttons: Abort, Retry, and Ignore.
Private Const MB_YESNOCANCEL       As Long = &H3&    '    The message box contains three push buttons: Yes, No, and Cancel.
Private Const MB_YESNO             As Long = &H4&    '    The message box contains two push buttons  : Yes and No.
Private Const MB_RETRYCANCEL       As Long = &H5&    '    The message box contains two push buttons  : Retry and Cancel.
Private Const MB_CANCELTRYCONTINUE As Long = &H6&    '    The message box contains three push buttons: Cancel, Try Again, Continue. Use this message box type instead of MB_ABORTRETRYIGNORE.
Private Const MB_HELP              As Long = &H4000& '    Adds a Help button to the message box. When the user clicks the Help button or presses F1, the system sends a WM_HELP message to the owner.
 

'Table 2
'To display an icon in the message box, specify one of the following values.
Private Const MB_ICONSTOP          As Long = &H10&   '    A stop-sign icon appears in the message box.
Private Const MB_ICONERROR         As Long = &H10&   '    A stop-sign icon appears in the message box.
Private Const MB_ICONHAND          As Long = &H10&   '    A stop-sign icon appears in the message box.
Private Const MB_ICONQUESTION      As Long = &H20&   '    A question-mark icon appears in the message box. The question-mark message icon is no longer recommended because it does not clearly represent a specific type of message and because the phrasing of a message as a question could apply to any message type. In addition, users can confuse the message symbol question mark with Help information. Therefore, do not use this question mark message symbol in your message boxes. The system continues to support its inclusion only for backward compatibility.
Private Const MB_ICONEXCLAMATION   As Long = &H30&   '    An exclamation-point icon appears in the message box.
Private Const MB_ICONWARNING       As Long = &H30&   '    An exclamation-point icon appears in the message box.
Private Const MB_ICONINFORMATION   As Long = &H40&   '    An icon consisting of a lowercase letter i in a circle appears in the message box.
Private Const MB_ICONASTERISK      As Long = &H40&   '    An icon consisting of a lowercase letter i in a circle appears in the message box.

'Table 3
'To indicate the default button, specify one of the following values.
Private Const MB_DEFBUTTON1        As Long = &H0&   '    The first button is the default button.
                                                    '    MB_DEFBUTTON1 is the default unless MB_DEFBUTTON2, MB_DEFBUTTON3, or MB_DEFBUTTON4 is specified.
Private Const MB_DEFBUTTON2        As Long = &H100& '    The second button is the default button.
Private Const MB_DEFBUTTON3        As Long = &H200& '    The third button is the default button.
Private Const MB_DEFBUTTON4        As Long = &H300& '    The fourth button is the default button.

'Table 4
'To indicate the modality of the dialog box, specify one of the following values.
Private Const MB_APPLMODAL         As Long = &H0&        '    The user must respond to the message box before continuing work in the window identified by the hWnd parameter. However, the user can move to the windows of other threads and work in those windows.
                                                         '    Depending on the hierarchy of windows in the application, the user may be able to move to other windows within the thread. All child windows of the parent of the message box are automatically disabled, but pop-up windows are not.
                                                         '    MB_APPLMODAL is the default if neither MB_SYSTEMMODAL nor MB_TASKMODAL is specified.
Private Const MB_SYSTEMMODAL       As Long = &H1000&     '    Same as MB_APPLMODAL except that the message box has the WS_EX_TOPMOST style. Use system-modal message boxes to notify the user of serious, potentially damaging errors that require immediate attention (for example, running out of memory). This flag has no effect on the user's ability to interact with windows other than those associated with hWnd.
Private Const MB_TASKMODAL         As Long = &H2000&     '    Same as MB_APPLMODAL except that all the top-level windows belonging to the current thread are disabled if the hWnd parameter is NULL. Use this flag when the calling application or library does not have a window handle available but still needs to prevent input to other windows in the calling thread without suspending other threads.
 

'Table 5
'To specify other options, use one or more of the following values.
Private Const MB_SETFOREGROUND        As Long = &H10000  '    The message box becomes the foreground window. Internally, the system calls the SetForegroundWindow function for the message box.
Private Const MB_DEFAULT_DESKTOP_ONLY As Long = &H20000  '    Same as desktop of the interactive window station. For more information, see Window Stations.
                                                         '    If the current input desktop is not the default desktop, MessageBox does not return until the user switches to the default desktop.
Private Const MB_TOPMOST              As Long = &H40000  '    The message box is created with the WS_EX_TOPMOST window style.
Private Const MB_RIGHT                As Long = &H80000  '    The text is right-justified.
Private Const MB_RTLREADING           As Long = &H100000 '    Displays message and caption text using right-to-left reading order on Hebrew and Arabic systems.
Private Const MB_SERVICE_NOTIFICATION As Long = &H200000 '    The caller is a service notifying the user of an event. The function displays a message box on the current active desktop, even if there is no user logged on to the computer.
                                                         '    Terminal Services: If the calling thread has an impersonation token, the function directs the message box to the session specified in the impersonation token.
                                                         '    If this flag is set, the hWnd parameter must be NULL. This is so that the message box can appear on a desktop other than the desktop corresponding to the hWnd.
                                                         '    For information on security considerations in regard to using this flag, see Interactive Services. In particular, be aware that this flag can produce interactive content on a locked desktop and should therefore be used for only a very limited set of scenarios, such as resource exhaustion.

Public Enum MsgBoxStyle
    
    'Button-style
    vbOKOnly = 0
    vbOKCancel = 1
    vbAbortRetryIgnore = 2
    vbYesNoCancel = 3
    vbYesNo = 4
    vbRetryCancel = 5
    vbCancelTryContinue = 6
    vbMsgBoxHelpButton = 16384    '(&H4000)
    
    'modality-style
    vbApplicationModal = 0
    vbSystemModal = &H1000& ' 4096 (&H1000)
    vbTaskModal = &H2000&
    
    'Icon-style
    vbCritical = 16    '(&H10)
    vbQuestion = 32    '(&H20)
    vbExclamation = 48 '(&H30)
    vbInformation = 64 '(&H40)
    
    'Default button when pressing Enter
    vbDefaultButton1 = 0
    vbDefaultButton2 = 256 '(&H100)
    vbDefaultButton3 = 512 '(&H200)
    vbDefaultButton4 = 768 '(&H300)
    
    vbMsgBoxSetForeground = 65536 '(&H10000)
    vbMsgBoxRight = 524288        '(&H80000)
    vbMsgBoxRtlReading = 1048576  '(&H100000)

End Enum

Public Enum MsgBoxFncType
    vbNormal = 0
    vbExtra = 1
    vbIndirect = 2
End Enum

'Maybe copy this to a module
#If Not Win64 Then
    
    Private Enum LongPtr
        [_]
    End Enum
    
#End If

Private Type MSGBOXPARAMS
    cbSize             As Long
    hwndOwner          As LongPtr 'Long
    hInstance          As LongPtr 'Long
    lpszText           As LongPtr 'String
    lpszCaption        As LongPtr 'String
    dwStyle            As Long    ' = wType = uType Type: UINT 'The contents and behavior of the dialog box. This parameter can be a combination of flags from the following groups of flags, see table 1-5
    lpszIcon           As LongPtr 'String  'lpszIcon Type: LPCTSTR
                                           'Identifies an icon resource. This parameter can be either a null-terminated string or an integer resource identifier passed to the MAKEINTRESOURCE macro.
                                           'To load one of the standard system-defined icons, set the hInstance member to NULL and set lpszIcon to one of the values listed with the LoadIcon function.
                                           'This member is ignored if the dwStyle member does not specify the MB_USERICON flag.
    dwContextHelpId    As Long
    lpfnMsgBoxCallback As LongPtr 'Type: MSGBOXCALLBACK
                                           'A pointer to the callback function that processes help events for the message box. The callback function has the following form:
                                           'VOID CALLBACK MsgBoxCallback(LPHELPINFO lpHelpInfo);
                                           'If this member is NULL, then the message box sends WM_HELP messages to the owner window when help events occur.
    dwLanguageId       As Long
End Type

Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" ( _
     ByVal hwnd As LongPtr, _
     ByRef lpText As Any, _
     ByRef lpCaption As Any, _
     ByVal wType As Long) As Long

Private Declare Function MessageBoxEx Lib "user32.dll" Alias "MessageBoxExW" ( _
     ByVal hwnd As LongPtr, _
     ByRef lpText As Any, _
     ByVal lpCaption As Any, _
     ByVal uType As Long, _
     ByVal wLanguageId As Long) As Long

Private Declare Function MessageBoxIndirect Lib "user32.dll" Alias "MessageBoxIndirectW" ( _
     ByRef lpMsgBoxParams As Any) As Long

Private m_MsgBoxFncType As MsgBoxFncType
Private m_MsgBoxParams  As MSGBOXPARAMS
Private m_Title         As String 'aka Caption
Private m_Text          As String 'aka Prompt
Private m_Result        As MsgBoxResult

Private Sub Class_Initialize()
    With m_MsgBoxParams
        .cbSize = LenB(m_MsgBoxParams)
        .dwStyle = vbOKOnly
        '.hwndOwner = ActiveWindow
        '.hInstance = App.hInstance
        m_Title = App.EXEName
    End With
End Sub

Public Property Get HIcon() As Long
    HIcon = m_MsgBoxParams.lpszIcon
End Property
Public Property Let HIcon(ByVal Value As Long)
    m_MsgBoxParams.lpszIcon = Value
End Property

Public Property Get Prompt() As String
    Prompt = m_Text 'aka Text
End Property
Public Property Let Prompt(ByVal Value As String)
    m_Text = Value 'aka Text
End Property
Public Property Get Text() As String
    Text = m_Text 'aka Prompt
End Property
Public Property Let Text(ByVal Value As String)
    m_Text = Value 'aka Prompt
End Property

Public Property Get Title() As String
    Title = m_Title 'aka Caption
End Property
Public Property Let Title(ByVal Value As String)
    m_Title = Value 'aka Caption
End Property
Public Property Get Caption() As String
    Caption = m_Title 'aka Title
End Property
Public Property Let Caption(ByVal Value As String)
    m_Title = Value 'aka Title
End Property

Public Property Get Style() As MsgBoxStyle
    Style = m_MsgBoxParams.dwStyle
End Property
Public Property Let Style(ByVal Value As MsgBoxStyle)
    m_MsgBoxParams.dwStyle = Value
End Property

Public Property Get MsgBoxFncType() As MsgBoxFncType
    MsgBoxFncType = m_MsgBoxFncType
End Property
Public Property Let MsgBoxFncType(ByVal Value As MsgBoxFncType)
    m_MsgBoxFncType = Value
End Property

Public Property Get LanguageID() As Long
    LanguageID = m_MsgBoxParams.dwLanguageId
End Property
Public Property Let LanguageID(ByVal Value As Long)
    m_MsgBoxParams.dwLanguageId = Value
End Property

Private Function FncPtr(pFnc As LongPtr) As LongPtr
    FncPtr = pFnc
End Function

Public Function Show(Optional aPrompt, Optional BtnsAndStyle, Optional aTitle, Optional aHelpFile, Optional aContext) As MsgBoxResult
    'Optional BtnsAndStyle As MsgBoxStyle = vbOKOnly
    If Not IsMissing(aPrompt) Then m_Text = CStr(aPrompt)
    If Not IsMissing(aTitle) Then m_Title = CStr(aTitle)
    With m_MsgBoxParams
        .lpfnMsgBoxCallback = FncPtr(AddressOf MWin.MessageBoxCallBack)
        .lpszCaption = StrPtr(m_Title)
        .lpszText = StrPtr(m_Text)
        If Not IsMissing(BtnsAndStyle) Then
            .dwStyle = CLng(BtnsAndStyle)
        End If
        '.lpszIcon = hIcon
        
        Select Case m_MsgBoxFncType
        Case 0: m_Result = MessageBox(.hwndOwner, ByVal .lpszText, ByVal .lpszCaption, .dwStyle)
        Case 1: m_Result = MessageBoxEx(.hwndOwner, ByVal .lpszText, ByVal .lpszCaption, .dwStyle, .dwLanguageId)
        Case 2: m_Result = MessageBoxIndirect(m_MsgBoxParams)
        End Select
        
    End With
    Show = m_Result
End Function

Public Function Result_ToStr(Optional aMsgBoxResult) As String
    Dim s As String
    Dim r As MsgBoxResult: r = IIf(IsMissing(aMsgBoxResult), m_Result, CLng(aMsgBoxResult))
    Select Case r
    Case vbOK:       s = "OK"
    Case vbCancel:   s = "Cancel"
    Case vbAbort:    s = "Abort"
    Case vbRetry:    s = "Retry"
    Case vbIgnore:   s = "Ignore"
    Case vbYes:      s = "Yes"
    Case vbNo:       s = "No"
    Case vbTryAgain: s = "TryAgain"
    Case vbContinue: s = "Continue"
    End Select
    Result_ToStr = s
End Function

'Function MyMsgbox(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As Variant, Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult
