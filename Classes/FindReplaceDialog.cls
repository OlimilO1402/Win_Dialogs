VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FindReplaceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://learn.microsoft.com/de-de/windows/win32/dlgbox/using-common-dialog-boxes#finding-text
'https://learn.microsoft.com/en-us/windows/win32/dlgbox/find-and-replace-dialog-boxes

Private Const CDERR_STRUCTSIZE      As Long = &H1    '    The lStructSize member of the initialization structure for the corresponding common dialog box is invalid.
Private Const CDERR_INITIALIZATION  As Long = &H2    '    The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available.
Private Const CDERR_NOTEMPLATE      As Long = &H3    '    The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template.
Private Const CDERR_NOHINSTANCE     As Long = &H4    '    The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle.
Private Const CDERR_LOADSTRFAILURE  As Long = &H5    '    The common dialog box function failed to load a specified string.
Private Const CDERR_FINDRESFAILURE  As Long = &H6    '    The common dialog box function failed to find a specified resource.
Private Const CDERR_LOADRESFAILURE  As Long = &H7    '    The common dialog box function failed to load a specified resource.
Private Const CDERR_LOCKRESFAILURE  As Long = &H8    '    The common dialog box function failed to lock a specified resource.
Private Const CDERR_MEMALLOCFAILURE As Long = &H9    '    The common dialog box function was unable to allocate memory for internal structures.
Private Const CDERR_MEMLOCKFAILURE  As Long = &HA    '    The common dialog box function was unable to lock the memory associated with a handle.
Private Const CDERR_NOHOOK          As Long = &HB    '    The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure.
Private Const CDERR_REGISTERMSGFAIL As Long = &HC    '    The RegisterWindowMessage function returned an error code when it was called by the common dialog box function.
Private Const CDERR_DIALOGFAILURE   As Long = &HFFFF& '    The dialog box could not be created. The common dialog box function's call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle.

'The following error codes can be returned for the PrintDlg function.
'Return code/value   Description

Private Const PDERR_CREATEICFAILURE  As Long = &H100A& '    The PrintDlg function failed when it attempted to create an information context.
Private Const PDERR_DEFAULTDIFFERENT As Long = &H100C& '    You called the PrintDlg function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer. This error occurs when you store the DEVNAMES structure, and the user changes the default printer by using the Control Panel.

'To use the printer described by the DEVNAMES structure, clear the DN_DEFAULTPRN flag and call PrintDlg again.
'To use the default printer, replace the DEVNAMES structure (and the structure, if one exists) with NULL; and call PrintDlg again.
Private Const PDERR_SETUPFAILURE     As Long = &H1001& '    The PrintDlg function failed to load the required resources.
Private Const PDERR_PARSEFAILURE     As Long = &H1002& '    The PrintDlg function failed to parse the strings in the [devices] section of the WIN.INI file.
Private Const PDERR_RETDEFFAILURE    As Long = &H1003& '    The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL.
Private Const PDERR_LOADDRVFAILURE   As Long = &H1004& '    The PrintDlg function failed to load the device driver for the specified printer.
Private Const PDERR_GETDEVMODEFAIL   As Long = &H1005& '    The printer driver failed to initialize a DEVMODE structure.
Private Const PDERR_INITFAILURE      As Long = &H1006& '    The PrintDlg function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function.
Private Const PDERR_NODEVICES        As Long = &H1007& '    No printer drivers were found.
Private Const PDERR_NODEFAULTPRN     As Long = &H1008& '    A default printer does not exist.
Private Const PDERR_DNDMMISMATCH     As Long = &H1009& '    The data in the DEVMODE and DEVNAMES structures describes two different printers.
Private Const PDERR_PRINTERNOTFOUND  As Long = &H100B& '    The [devices] section of the WIN.INI file did not contain an entry for the requested printer.

'The following error codes can be returned for the ChooseFont function.
'Return code/value   Description
Private Const CFERR_NOFONTS          As Long = &H2001& '    No fonts exist.
Private Const CFERR_MAXLESSTHANMIN   As Long = &H2002& '    The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member.

'*The following error codes can be returned for the GetOpenFileName and GetSaveFileName functions.
'return code/value   Description
Private Const FNERR_SUBCLASSFAILURE  As Long = &H3001& '    An attempt to subclass a list box failed because sufficient memory was not available.
Private Const FNERR_INVALIDFILENAME  As Long = &H3002& '    A file name is invalid.
Private Const FNERR_BUFFERTOOSMALL   As Long = &H3003& '    The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the file name specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size required to receive the full name, in characters.

'The following error code can be returned for the FindText and ReplaceText functions.
'Return code/value   Description
Private Const FRERR_BUFFERLENGTHZERO As Long = &H4001& '   A member of the FINDREPLACE structure points to an invalid buffer.





Private Const FR_DOWN                 As Long = &H1       '    If set, the Down button of the direction radio buttons in a Find dialog box is selected indicating that you should search from the current location to the end of the document. If not set, the Up button is selected so you should search to the beginning of the document. You can set this flag to initialize the dialog box. If set in a FINDMSGSTRING message, indicates the user's selection.
Private Const FR_WHOLEWORD            As Long = &H2       '    If set, the Match Whole Word Only check box is selected indicating that you should search only for whole words that match the search string. If not set, the check box is unselected so you should also search for word fragments that match the search string. You can set this flag to initialize the dialog box. If set in a FINDMSGSTRING message, indicates the user's selection.
Private Const FR_MATCHCASE            As Long = &H4       '    If set, the Match Case check box is selected indicating that the search should be case-sensitive. If not set, the check box is unselected so the search should be case-insensitive. You can set this flag to initialize the dialog box. If set in a FINDMSGSTRING message, indicates the user's selection.
Private Const FR_FINDNEXT             As Long = &H8       '    If set in a FINDMSGSTRING message, indicates that the user clicked the Find Next button in a Find or Replace dialog box. The lpstrFindWhat member specifies the string to search for.
Private Const FR_REPLACE              As Long = &H10      '    If set in a FINDMSGSTRING message, indicates that the user clicked the Replace button in a Replace dialog box. The lpstrFindWhat member specifies the string to be replaced and the lpstrReplaceWith member specifies the replacement string.
Private Const FR_REPLACEALL           As Long = &H20      '    If set in a FINDMSGSTRING message, indicates that the user clicked the Replace All button in a Replace dialog box. The lpstrFindWhat member specifies the string to be replaced and the lpstrReplaceWith member specifies the replacement string.
Private Const FR_DIALOGTERM           As Long = &H40      '    If set in a FINDMSGSTRING message, indicates that the dialog box is closing. When you receive a message with this flag set, the dialog box handle returned by the FindText or ReplaceText function is no longer valid.
Private Const FR_SHOWHELP             As Long = &H80      '    Causes the dialog box to display the Help button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the Help button.
Private Const FR_ENABLEHOOK           As Long = &H100     '    Enables the hook function specified in the lpfnHook member. This flag is used only to initialize the dialog box.
Private Const FR_ENABLETEMPLATE       As Long = &H200     '    Indicates that the hInstance and lpTemplateName members specify a dialog box template to use in place of the default template. This flag is used only to initialize the dialog box.
Private Const FR_NOUPDOWN             As Long = &H400     '    If set when initializing a Find dialog box, disables the search direction radio buttons.
Private Const FR_NOMATCHCASE          As Long = &H800     '    If set when initializing a Find or Replace dialog box, disables the Match Case check box.
Private Const FR_NOWHOLEWORD          As Long = &H1000    '    If set when initializing a Find or Replace dialog box, disables the Whole Word check box.
Private Const FR_ENABLETEMPLATEHANDLE As Long = &H2000    '    Indicates that the hInstance member identifies a data block that contains a preloaded dialog box template. The system ignores the lpTemplateName member if this flag is specified.
Private Const FR_HIDEUPDOWN           As Long = &H4000    '    If set when initializing a Find dialog box, hides the search direction radio buttons.
Private Const FR_HIDEMATCHCASE        As Long = &H8000    '    If set when initializing a Find or Replace dialog box, hides the Match Case check box.
Private Const FR_HIDEWHOLEWORD        As Long = &H10000   '    If set when initializing a Find or Replace dialog box, hides the Match Whole Word Only check box.

'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/ns-commdlg-findreplacew
'typedef struct tagFINDREPLACEW {
'  DWORD        lStructSize;
'  HWND         hwndOwner;
'  HINSTANCE    hInstance;
'  DWORD        Flags;
'  LPWSTR       lpstrFindWhat;
'  LPWSTR       lpstrReplaceWith;
'  WORD         wFindWhatLen;
'  WORD         wReplaceWithLen;
'  LPARAM       lCustData;
'  LPFRHOOKPROC lpfnHook;
'  LPCWSTR      lpTemplateName;
'} FINDREPLACEW, *LPFINDREPLACEW;
Private Type FINDREPLACEW
    lStructSize      As Long
    hwndOwner        As LongPtr
    hInstance        As LongPtr
    flags            As Long
    lpstrFindWhat    As LongPtr 'String  'LPWSTR
    lpstrReplaceWith As LongPtr 'String  'LPWSTR
    wFindWhatLen     As Integer
    wReplaceWithLen  As Integer
    lCustData        As Long
    lpfnHook         As LongPtr 'LPFRHOOKPROC
    lpTemplateName   As LongPtr 'String  'LPCWSTR
End Type ' FINDREPLACEW, *LPFINDREPLACEW

'LPFRHOOKPROC Lpfrhookproc;
'UINT_PTR Lpfrhookproc(
'  HWND unnamedParam1,
'  UINT unnamedParam2,
'  WPARAM unnamedParam3,
'  lParam unnamedParam4
')
'{...}

#If VBA7 Then
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-findtextw
    'HWND FindTextW(  [in] LPFINDREPLACEW unnamedParam1 );
    Private Declare PtrSafe Function FindTextW Lib "comdlg32" (pFindreplace As FINDREPLACEW) As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-replacetextw
    'HWND ReplaceTextW(  [in, out] LPFINDREPLACEW unnamedParam1 );
    Private Declare PtrSafe Function ReplaceTextW Lib "comdlg32.dll" (pFindreplace As FINDREPLACEW) As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-commdlgextendederror
    Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long
    
#Else
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-findtextw
    'HWND FindTextW(  [in] LPFINDREPLACEW unnamedParam1 );
    Private Declare Function FindTextW Lib "comdlg32" (pFindreplace As FINDREPLACEW) As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-replacetextw
    'HWND ReplaceTextW(  [in, out] LPFINDREPLACEW unnamedParam1 );
    Private Declare Function ReplaceTextW Lib "comdlg32" (pFindreplace As FINDREPLACEW) As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-commdlgextendederror
    Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
    
#End If

Private m_FindReplaceDlg As FINDREPLACEW
Private mFlags           As Long
Private m_hWnd           As LongPtr
Private m_FindWhat       As String
Private m_ReplaceWith    As String
Private m_IsReplaceDlg   As Boolean
Private m_Err            As Long

Private Sub Class_Initialize()
    With m_FindReplaceDlg
        .lStructSize = LenB(m_FindReplaceDlg)
        .hInstance = App.hInstance
        '.lpfnHook = FncPtr(AddressOf MCallBack.FindReplaceCallBack)
    End With
    mFlags = mFlags Or FR_DOWN
End Sub

Friend Sub New_(ByVal FindWhat As String, Optional ByVal ReplaceWith As String, Optional ByVal bReplaceDlg As Boolean = False)
    m_FindWhat = FindWhat: m_ReplaceWith = ReplaceWith: m_IsReplaceDlg = bReplaceDlg
End Sub

Public Sub ShowDialog(Optional Frm As Object)
    With m_FindReplaceDlg
        .flags = mFlags
        If Not IsMissing(Frm) Then
            .hwndOwner = Frm.hwnd
        End If
        .lpstrFindWhat = StrPtr(m_FindWhat & vbNullChar)
        .wFindWhatLen = Len(m_FindWhat & vbNullChar)
        .lpstrReplaceWith = StrPtr(m_ReplaceWith & vbNullChar)
        .wReplaceWithLen = Len(m_ReplaceWith & vbNullChar)
    End With
    m_hWnd = FindTextW(m_FindReplaceDlg)
    
    If m_hWnd = 0 Then
        m_Err = CommDlgExtendedError
    End If
End Sub

Public Property Get IsReplaceDlg() As Boolean
    IsReplaceDlg = m_IsReplaceDlg
End Property
Public Property Let IsReplaceDlg(ByVal Value As Boolean)
    m_IsReplaceDlg = Value
End Property

Public Property Get FindWhat() As String
    FindWhat = m_FindWhat
End Property
Public Property Let FindWhat(ByVal Value As String)
    m_FindWhat = Value
End Property

Public Property Get ReplaceWith() As String
    ReplaceWith = m_ReplaceWith
End Property
Public Property Let ReplaceWith(ByVal Value As String)
    m_ReplaceWith = Value
End Property

Public Property Get SearchDown() As Boolean
    SearchDown = mFlags And FR_DOWN
End Property
Public Property Let SearchDown(ByVal Value As Boolean)
    mFlags = mFlags Or FR_DOWN
    If Not Value Then mFlags = mFlags Xor FR_DOWN
End Property

Public Property Get MatchWholeWordOnly() As Boolean
    MatchWholeWordOnly = mFlags And FR_WHOLEWORD
End Property
Public Property Let MatchWholeWordOnly(ByVal Value As Boolean)
    mFlags = mFlags Or FR_WHOLEWORD
    If Not Value Then mFlags = mFlags Xor FR_WHOLEWORD
End Property

Public Property Get MatchCase() As Boolean
    MatchCase = mFlags And FR_MATCHCASE
End Property
Public Property Let MatchCase(ByVal Value As Boolean)
    mFlags = mFlags Or FR_MATCHCASE
    If Not Value Then mFlags = mFlags Xor FR_MATCHCASE
End Property

Public Property Get FindNext() As Boolean
    FindNext = mFlags And FR_FINDNEXT
End Property
Public Property Let FindNext(ByVal Value As Boolean)
    mFlags = mFlags Or FR_FINDNEXT
    If Not Value Then mFlags = mFlags Xor FR_FINDNEXT
End Property

Public Property Get Replace() As Boolean
    Replace = mFlags And FR_REPLACE
End Property
Public Property Let Replace(ByVal Value As Boolean)
    mFlags = mFlags Or FR_REPLACE
    If Not Value Then mFlags = mFlags Xor FR_REPLACE
End Property

Public Property Get ReplaceAll() As Boolean
    ReplaceAll = mFlags And FR_REPLACEALL
End Property
Public Property Let ReplaceAll(ByVal Value As Boolean)
    mFlags = mFlags Or FR_REPLACEALL
    If Not Value Then mFlags = mFlags Xor FR_REPLACEALL
End Property

Public Property Get IsTerminating() As Boolean
    IsTerminating = mFlags And FR_DIALOGTERM
End Property
Public Property Let IsTerminating(ByVal Value As Boolean)
    mFlags = mFlags Or FR_DIALOGTERM
    If Not Value Then mFlags = mFlags Xor FR_DIALOGTERM
End Property

Public Property Get ShowHelp() As Boolean
    ShowHelp = mFlags And FR_SHOWHELP
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    mFlags = mFlags Or FR_SHOWHELP
    If Not Value Then mFlags = mFlags Xor FR_SHOWHELP
End Property

Public Property Get NoUpDown() As Boolean
    NoUpDown = mFlags And FR_NOUPDOWN
End Property
Public Property Let NoUpDown(ByVal Value As Boolean)
    mFlags = mFlags Or FR_NOUPDOWN
    If Not Value Then mFlags = mFlags Xor FR_NOUPDOWN
End Property

Public Property Get NoMatchCase() As Boolean
    NoMatchCase = mFlags And FR_NOMATCHCASE
End Property
Public Property Let NoMatchCase(ByVal Value As Boolean)
    mFlags = mFlags Or FR_NOMATCHCASE
    If Not Value Then mFlags = mFlags Xor FR_NOMATCHCASE
End Property

Public Property Get NoWholeWord() As Boolean
    NoWholeWord = mFlags And FR_NOWHOLEWORD
End Property
Public Property Let NoWholeWord(ByVal Value As Boolean)
    mFlags = mFlags Or FR_NOWHOLEWORD
    If Not Value Then mFlags = mFlags Xor FR_NOWHOLEWORD
End Property

Public Property Get HideUpDown() As Boolean
    HideUpDown = mFlags And FR_HIDEUPDOWN
End Property
Public Property Let HideUpDown(ByVal Value As Boolean)
    mFlags = mFlags Or FR_HIDEUPDOWN
    If Not Value Then mFlags = mFlags Xor FR_HIDEUPDOWN
End Property

Public Property Get HideWholeWord() As Boolean
    HideWholeWord = mFlags And FR_HIDEWHOLEWORD
End Property
Public Property Let HideWholeWord(ByVal Value As Boolean)
    mFlags = mFlags Or FR_HIDEWHOLEWORD
    If Not Value Then mFlags = mFlags Xor FR_HIDEWHOLEWORD
End Property

Public Property Get HideMatchCase() As Boolean
    HideMatchCase = mFlags And FR_HIDEMATCHCASE
End Property
Public Property Let HideMatchCase(ByVal Value As Boolean)
    mFlags = mFlags Or FR_HIDEMATCHCASE
    If Not Value Then mFlags = mFlags Xor FR_HIDEMATCHCASE
End Property

Public Property Get LastError() As String
    Dim s As String
    Select Case m_Err
    Case CDERR_STRUCTSIZE: s = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
    Case CDERR_INITIALIZATION: s = "The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available."
    Case CDERR_NOTEMPLATE: s = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
    Case CDERR_NOHINSTANCE: s = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
    Case CDERR_LOADSTRFAILURE: s = "The common dialog box function failed to load a specified string."
    Case CDERR_FINDRESFAILURE: s = "The common dialog box function failed to find a specified resource."
    Case CDERR_LOADRESFAILURE: s = "The common dialog box function failed to load a specified resource."
    Case CDERR_LOCKRESFAILURE: s = "The common dialog box function failed to lock a specified resource."
    Case CDERR_MEMALLOCFAILURE: s = "The common dialog box function was unable to allocate memory for internal structures."
    Case CDERR_MEMLOCKFAILURE: s = "The common dialog box function was unable to lock the memory associated with a handle."
    Case CDERR_NOHOOK: s = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
    Case CDERR_REGISTERMSGFAIL: s = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
    Case CDERR_DIALOGFAILURE: s = "The dialog box could not be created. The common dialog box function's call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle."

'The following error codes can be returned for the PrintDlg function.
'Return code/value   Description

    Case PDERR_CREATEICFAILURE: s = "The PrintDlg function failed when it attempted to create an information context."
    Case PDERR_DEFAULTDIFFERENT: s = "You called the PrintDlg function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer. This error occurs when you store the DEVNAMES structure, and the user changes the default printer by using the Control Panel."

'To use the printer described by the DEVNAMES structure, clear the DN_DEFAULTPRN flag and call PrintDlg again.
'To use the default printer, replace the DEVNAMES structure (and the structure, if one exists) with NULL; and call PrintDlg again.
    Case PDERR_SETUPFAILURE: s = "The PrintDlg function failed to load the required resources."
    Case PDERR_PARSEFAILURE: s = "The PrintDlg function failed to parse the strings in the [devices] section of the WIN.INI file."
    Case PDERR_RETDEFFAILURE: s = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL."
    Case PDERR_LOADDRVFAILURE: s = "The PrintDlg function failed to load the device driver for the specified printer."
    Case PDERR_GETDEVMODEFAIL: s = "The printer driver failed to initialize a DEVMODE structure."
    Case PDERR_INITFAILURE: s = "The PrintDlg function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function."
    Case PDERR_NODEVICES: s = "No printer drivers were found."
    Case PDERR_NODEFAULTPRN: s = "A default printer does not exist."
    Case PDERR_DNDMMISMATCH: s = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
    Case PDERR_PRINTERNOTFOUND: s = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."

'The following error codes can be returned for the ChooseFont function.
'Return code/value   Description
    Case CFERR_NOFONTS: s = "No fonts exist."
    Case CFERR_MAXLESSTHANMIN: s = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."

'*The following error codes can be returned for the GetOpenFileName and GetSaveFileName functions.
'return code/value   Description
    Case FNERR_SUBCLASSFAILURE: s = "An attempt to subclass a list box failed because sufficient memory was not available."
    Case FNERR_INVALIDFILENAME: s = "A file name is invalid."
    Case FNERR_BUFFERTOOSMALL: s = "The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the file name specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size required to receive the full name, in characters."

'The following error code can be returned for the FindText and ReplaceText functions.
'Return code/value   Description
    Case FRERR_BUFFERLENGTHZERO: s = "A member of the FINDREPLACE structure points to an invalid buffer."
    
    End Select
    LastError = s
End Property





'Suchen nach Text
'
'In diesem Thema wird Beispielcode beschrieben, der ein Dialogfeld Suchen anzeigt und verwaltet, sodass der Benutzer die Parameter eines
'Suchvorgangs angeben kann. Das Dialogfeld sendet Nachrichten an die Fensterprozedur, damit Sie den Suchvorgang ausführen können.
'
'Der Code zum Anzeigen und Verwalten eines Dialogfelds Ersetzen ist ähnlich, mit der Ausnahme, dass es die ReplaceText-Funktion verwendet,
'um das Dialogfeld anzuzeigen. Das Dialogfeld Ersetzen sendet auch Nachrichten als Reaktion auf Benutzerklicks auf die Schaltflächen
'Ersetzen und Alle ersetzen .
'
'Um das Dialogfeld Suchen oder Ersetzen verwenden zu können, müssen Sie drei separate Aufgaben ausführen:
'
'  *  Rufen Sie einen Nachrichtenbezeichner für die registrierte FINDMSGSTRING-Nachricht ab.
'  *  Zeigt das Dialogfeld an.
'  *  Verarbeiten sie FINDMSGSTRING-Nachrichten , wenn das Dialogfeld geöffnet ist.
'
'Wenn Sie Ihre Anwendung initialisieren, rufen Sie die Funktion RegisterWindowMessage auf, um einen Nachrichtenbezeichner für die
'registrierte FINDMSGSTRING-Nachricht abzurufen.
'
'UINT uFindReplaceMsg;  // message identifier for FINDMSGSTRING
'
'uFindReplaceMsg = RegisterWindowMessage(FINDMSGSTRING);
'
'Um ein Suchdialogfeld anzuzeigen, initialisieren Sie zuerst eine FINDREPLACE-Struktur , und rufen Sie dann die FindText-Funktion auf.
'Beachten Sie, dass die FINDREPLACE-Struktur und der Puffer für die Suchzeichenfolge eine globale oder statische Variable sein sollten,
'damit sie nicht aus dem Bereich herauskommt, bevor das Dialogfeld geschlossen wird.
'Sie müssen das hwndOwner-Element festlegen, um das Fenster anzugeben, in dem die registrierten Nachrichten empfangen werden.
'Nachdem Sie das Dialogfeld erstellt haben, können Sie es mithilfe des zurückgegebenen Handle verschieben oder bearbeiten.
'
'FINDREPLACE fr;       // common dialog box structure
'HWND hwnd;            // owner window
'CHAR szFindWhat[80];  // buffer receiving string
'HWND hdlg = NULL;     // handle to Find dialog box
'
'// Initialize FINDREPLACE
'ZeroMemory(&fr, sizeof(fr));
'fr.lStructSize = sizeof(fr);
'fr.hwndOwner = hwnd;
'fr.lpstrFindWhat = szFindWhat;
'fr.wFindWhatLen = 80;
'fr.Flags = 0;
'
'hdlg = FindText(&fr);
'
'Wenn das Dialogfeld geöffnet ist, muss Ihre Standard Nachrichtenschleife einen Aufruf der IsDialogMessage-Funktion enthalten.
'Übergeben Sie ein Handle als Parameter im IsDialogMessage-Aufruf an das Dialogfeld.
'Dadurch wird sichergestellt, dass das Dialogfeld Tastaturmeldungen ordnungsgemäß verarbeitet.
'
'Zum Überwachen von Nachrichten, die aus dem Dialogfeld gesendet werden, muss ihre Fensterprozedur nach der registrierten
'FINDMSGSTRING-Nachricht suchen und die werte verarbeiten, die in der FINDREPLACE-Struktur übergeben werden, wie im folgenden Beispiel
'beschrieben.
'
'LPFINDREPLACE lpfr;
'
'if (message == uFindReplaceMsg)
'{
'    // Get pointer to FINDREPLACE structure from lParam.
'    lpfr = (LPFINDREPLACE)lParam;
'
'    // If the FR_DIALOGTERM flag is set,
'    // invalidate the handle that identifies the dialog box.
'    if (lpfr->Flags & FR_DIALOGTERM)
'    {
'        hdlg = NULL;
'        return 0;
'    }
'
'    // If the FR_FINDNEXT flag is set,
'    // call the application-defined search routine
'    // to search for the requested string.
'    if (lpfr->Flags & FR_FINDNEXT)
'    {
'        SearchFile(lpfr->lpstrFindWhat,
'                   (BOOL) (lpfr->Flags & FR_DOWN),
'                   (BOOL) (lpfr->Flags & FR_MATCHCASE));
'    }
'
'    return 0;
'}
