VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderBrowserDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class FolderBrowserDialog
'          Inherits System.Windows.Forms.CommonDialog
'     Member von: System.Windows.Forms
'
'Zusammenfassung:
' Stellt ein Standarddialogfeld dar, in dem Benutzer einen Ordner auswählen können.
Option Explicit
Implements ICallBack
'Private Const CSIDL_DESKTOP                 As Long = &H0
'Private Const CSIDL_INTERNET                As Long = &H1
'Private Const CSIDL_PROGRAMS                As Long = &H2
'Private Const CSIDL_CONTROLS                As Long = &H3
'Private Const CSIDL_PRINTERS                As Long = &H4
'Private Const CSIDL_PERSONAL                As Long = &H5
'Private Const CSIDL_FAVORITES               As Long = &H6
'Private Const CSIDL_STARTUP                 As Long = &H7
'Private Const CSIDL_RECENT                  As Long = &H8
'Private Const CSIDL_SENDTO                  As Long = &H9
'Private Const CSIDL_BITBUCKET               As Long = &HA
'Private Const CSIDL_STARTMENU               As Long = &HB
''&HC, &HD, &HE, &HF ??
'
'Private Const CSIDL_DESKTOPDIRECTORY        As Long = &H10
'Private Const CSIDL_DRIVES                  As Long = &H11
'Private Const CSIDL_NETWORK                 As Long = &H12
'Private Const CSIDL_NETHOOD                 As Long = &H13
'Private Const CSIDL_FONTS                   As Long = &H14
'Private Const CSIDL_TEMPLATES               As Long = &H15
'Private Const CSIDL_COMMON_STARTMENU        As Long = &H16
'Private Const CSIDL_COMMON_PROGRAMS         As Long = &H17
'Private Const CSIDL_COMMON_STARTUP          As Long = &H18
'Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
'Private Const CSIDL_APPDATA                 As Long = &H1A
'Private Const CSIDL_PRINTHOOD               As Long = &H1B
'Private Const CSIDL_LOCAL_APPDATA           As Long = &H1C
'Private Const CSIDL_ALTSTARTUP              As Long = &H1D
'Private Const CSIDL_COMMON_ALTSTARTUP       As Long = &H1E
'Private Const CSIDL_COMMON_FAVORITES        As Long = &H1F
'
'Private Const CSIDL_INTERNET_CACHE          As Long = &H20
'Private Const CSIDL_COOKIES                 As Long = &H21
'Private Const CSIDL_HISTORY                 As Long = &H22
'Private Const CSIDL_COMMON_APPDATA          As Long = &H23
'Private Const CSIDL_WINDOWS                 As Long = &H24
'Private Const CSIDL_SYSTEM                  As Long = &H25
'Private Const CSIDL_PROGRAM_FILES           As Long = &H26
'Private Const CSIDL_MYPICTURES              As Long = &H27
'Private Const CSIDL_PROFILE                 As Long = &H28
'Private Const CSIDL_SYSTEMX86               As Long = &H29
'Private Const CSIDL_PROGRAM_FILESX86        As Long = &H2A
'Private Const CSIDL_PROGRAM_FILES_COMMON    As Long = &H2B
'Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
'Private Const CSIDL_COMMON_TEMPLATES        As Long = &H2D
'Private Const CSIDL_COMMON_DOCUMENTS        As Long = &H2E
'Private Const CSIDL_COMMON_ADMINTOOLS       As Long = &H2F
'
'Private Const CSIDL_ADMINTOOLS              As Long = &H30
'Private Const CSIDL_CONNECTIONS             As Long = &H31
'Private Const CSIDL_FLAG_DONT_VERIFY        As Long = &H4000
'
'Private Const CSIDL_FLAG_CREATE As Long = &H8000
'Private Const CSIDL_FLAG_MASK As Long = &HFF00&
'Private Const CSIDL_FLAG_PFTI_TRACKTARGET As Long = CSIDL_FLAG_DONT_VERIFY
'
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
'im .NET-FX gibt es auch ein Enum Environment.SpecialFolder
Public Enum Environment_SpecialFolder
    SpecialFolder_Desktop = &H0    ' = CSIDL_DESKTOP
    CSIDL_INTERNET = &H1
    SpecialFolder_Programs = &H2   ' = CSIDL_PROGRAMS 'Programmgruppen im Startverzeichnis
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    SpecialFolder_Personal = &H5   ' = CSIDL_PERSONAL
    SpecialFolder_Favorites = &H6  ' = CSIDL_FAVORITES
    SpecialFolder_Startup = &H7    ' = CSIDL_STARTUP
    SpecialFolder_Recent = &H8     ' = CSIDL_RECENT
    SpecialFolder_SendTo = &H9     ' = CSIDL_SENDTO
    CSIDL_BITBUCKET = &HA          'Papierkorb
    SpecialFolder_StartMenu = &HB  ' = CSIDL_STARTMENU
    '&HC ??
    SpecialFolder_MyMusic = &HD
    '&HE, &HF ??
    SpecialFolder_DesktopDirectory = &H10
    SpecialFolder_MyComputer = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    SpecialFolder_Templates = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    SpecialFolder_ApplicationData = &H1A
    CSIDL_PRINTHOOD = &H1B
    SpecialFolder_LocalApplicationData = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    SpecialFolder_InternetCache = &H20
    SpecialFolder_Cookies = &H21
    SpecialFolder_History = &H22
    SpecialFolder_CommonApplicationData = &H23
    CSIDL_WINDOWS = &H24
    SpecialFolder_System = &H25
    SpecialFolder_ProgramFiles = &H26 'Programmdateien
    SpecialFolder_MyPictures = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    SpecialFolder_CommonProgramFiles = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00&
    CSIDL_FLAG_PFTI_TRACKTARGET = CSIDL_FLAG_DONT_VERIFY
End Enum

Public Enum DialogResult
    DialogResult_None = 0
    DialogResult_OK = VbMsgBoxResult.vbOK
    DialogResult_Cancel = VbMsgBoxResult.vbCancel
    DialogResult_Abort = VbMsgBoxResult.vbAbort
    DialogResult_Retry = VbMsgBoxResult.vbRetry
    DialogResult_Ignore = VbMsgBoxResult.vbIgnore
    DialogResult_Yes = VbMsgBoxResult.vbYes
    DialogResult_No = VbMsgBoxResult.vbNo
End Enum

Private Type BrowseInfoW
    hwndOwner      As LongPtr
    pidlRoot       As Long
    pszDisplayName As LongPtr 'String
    lpszTitle      As LongPtr 'String
    ulFlags        As Long
    lpfn           As LongPtr
    lParam         As Long
    iImage         As Long
End Type

Private mFlags        As Long
Private mDescription  As String
Private mRootFolder   As Long
Private mSelectedPath As String
Private mTag          As Variant

Private Const WM_USER              As Long = &H400
Public Enum BrowseInfoFlags 'ulFlags:
    BIF_RETURNONLYFSDIRS = &H1        ': Gestattet nur Dateisystemordner als Auswahl.
    BIF_DONTGOBELOWDOMAIN = &H2       ': Der Dialog zeigt keine Netzwerkordner unterhalb der aktuellen Domain.
    BIF_STATUSTEXT = &H4              ': Der Dialog enthält eine Statuszeile. Die Rückruffunktion kann die Statuszeile ausfüllen.
    BIF_RETURNFSANCESTORS = &H8       ': Gestattet nur Dateisystemobjekte als Auswahl
    BIF_EDITBOX = &H10
    BIF_VALIDATE = &H20
    BIF_NEWDIALOGSTYLE = &H40
    BIF_USENEWUI = &H40
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_DONTSHOWNEWFOLDERBUTTON = &H200  '512
    BFFM_SETSTATUSTEXTA = (WM_USER + 100)
    BFFM_ENABLEOK = (WM_USER + 101)      '1125
    BFFM_SETSELECTIONA = (WM_USER + 102) '1126
    BFFM_SETSELECTIONW = (WM_USER + 103)
    BFFM_SETSTATUSTEXTW = (WM_USER + 104)
    BIF_BROWSEFORCOMPUTER = &H1000       ': Als Auswahl sind nur Computer erlaubt. Wenn der Anwender andere Objekte, also Ordner oder Laufwerke markiert, kann der OK-Button nicht ausgewählt werden.
    BIF_BROWSEFORPRINTER = &H2000        ': Gestattet nur Drucker als Auswahl.
    BIF_BROWSEINCLUDEFILES = &H4000      ': Der Dialog zeigt neben Computern, Laufwerken und Ordnern auch Dateien an.
    BIF_SHAREABLE = &H8000
End Enum
Private Const BFFM_INITIALIZED     As Long = 1
Private Const BFFM_SELCHANGED      As Long = 2
'Private Const BFFM_VALIDATEFAILEDA As Long = 3
'Private Const BFFM_VALIDATEFAILEDW As Long = 4
'
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As LongPtr, ByVal csidl As Long, ByRef ppidl_out As Long) As Long      'HRESULT
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (pBrowseInfo As BrowseInfoW) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As LongPtr) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetActiveWindow Lib "user32" () As LongPtr

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByRef pv As Any)

Public Event HelpRequest(ByVal sender As FolderBrowserDialog)

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Property Get Description() As String
' Ruft den beschreibenden Text ab, der im Dialogfeld über dem
' Strukturansichts-Steuerelement angezeigt wird, oder legt diesen fest.
    Description = mDescription
End Property
Public Property Let Description(StrVal As String)
    mDescription = StrVal
End Property

Public Property Get Flags() As BrowseInfoFlags
    Flags = mFlags
End Property
Public Property Let Flags(FVal As BrowseInfoFlags)
    mFlags = FVal
End Property

Public Sub Reset()
    mFlags = 0&
    'mFlags = mFlags Or BIF_RETURNONLYFSDIRS
    'mFlags = mFlags Or BIF_DONTGOBELOWDOMAIN
    'mFlags = mFlags Or BIF_STATUSTEXT
    'mFlags = mFlags Or BIF_RETURNFSANCESTORS
    
    'mFlags = mFlags Or BIF_EDITBOX
    'mFlags = mFlags Or BIF_VALIDATE
    mFlags = mFlags Or BIF_NEWDIALOGSTYLE
    mFlags = mFlags Or BIF_USENEWUI
    'mFlags = mFlags Or BIF_BROWSEINCLUDEURLS
    
    'om 2007_02_06 neue Const
    'mFlags = mFlags Or BIF_DONTSHOWNEWFOLDERBUTTON
    
    'mFlags = mFlags Or BIF_BROWSEFORCOMPUTER
    'mFlags = mFlags Or BIF_BROWSEFORPRINTER
    'mFlags = mFlags Or BIF_BROWSEINCLUDEFILES
    'mFlags = mFlags Or BIF_SHAREABLE
End Sub

Public Property Let RootFolder(LngVal As Environment_SpecialFolder)
    mRootFolder = LngVal
End Property
Public Property Get RootFolder() As Environment_SpecialFolder
    ' Ruft den Stammordner ab, von dem aus eine Suche gestartet wird, oder legt diesen fest.
    RootFolder = mRootFolder
End Property

Public Property Let SelectedPath(StrVal As String)
    mSelectedPath = StrVal
End Property
Public Property Get SelectedPath() As String
    ' Ruft den von den Benutzern ausgewählten Pfad ab oder legt diesen fest.
    SelectedPath = mSelectedPath
End Property

Public Property Let ShowNewFolderButton(BolVal As Boolean)
    mFlags = mFlags Or BIF_DONTSHOWNEWFOLDERBUTTON
    If BolVal Then
        mFlags = mFlags Xor BIF_DONTSHOWNEWFOLDERBUTTON
    End If
End Property
Public Property Get ShowNewFolderButton() As Boolean
    ' Ruft den Wert ab, der angibt, ob die Schaltfläche New Folder im Dialogfeld für die
    ' Ordnersuche angezeigt wird.
    ShowNewFolderButton = Not (mFlags And BIF_DONTSHOWNEWFOLDERBUTTON)
End Property

'wie schon hinlänglich bekann, füg einfach irgendwas hinzu
'wird intern nicht verwendet, kann extern verwendet werden
'Public Property Get Tag() As Object
Public Property Get Tag() As Variant
    Tag = mTag
End Property
Public Property Let Tag(VarVal As Variant)
    mTag = VarVal
End Property

Public Function ToString() As String
    ToString = "Windows.Forms.FolderBrowserDialog"
End Function

Public Function ShowDialog(Optional Frm As Variant) As DialogResult
    Dim BI         As BrowseInfoW
    Dim hhwndOwner As LongPtr
  
    If IsMissing(Frm) Then
        hhwndOwner = GetActiveWindow
    Else
        hhwndOwner = Frm.hwnd
    End If
    
Try: On Error GoTo Catch
    Dim pRoot      As Long
    Call SHGetSpecialFolderLocation(hhwndOwner, mRootFolder, pRoot)
    If (pRoot = 0&) Then
        Call SHGetSpecialFolderLocation(hhwndOwner, 0, pRoot)
        If (pRoot = 0&) Then
            'Throw New InvalidOperationException(SR.GetString("FolderBrowserDialogNoRootFolder"))
            MsgBox "FolderBrowserDialogNoRootFolder"
            Exit Function
        End If
    End If
    Dim descr As String: descr = mDescription & vbNullChar
    Dim pDescr As LongPtr: pDescr = StrPtr(descr)
    With BI
        .hwndOwner = hhwndOwner
        .pidlRoot = pRoot
        '.pszDisplayName = ""
        'Beschreibung im Dialog 'als nullterminierter String
        .lpszTitle = pDescr 'mDescription  '& vbNullChar
        .ulFlags = mFlags
        .lpfn = FncPtr(AddressOf ModCallBack.FolderBrowserDialogCallBack)
        .lParam = ObjPtr(Me)
        .iImage = 0&
    End With
    
    'Anzeigen des Dialogs und Übergabe an eine IID-Liste
    Dim IDList As Long
    IDList = SHBrowseForFolder(BI)
    If IDList <> 0 Then
        Dim Buffer As String: Buffer = String$(1024, vbNullChar)
        Dim pBuffer As LongPtr: pBuffer = StrPtr(pBuffer)
        Call SHGetPathFromIDList(IDList, ByVal pBuffer)
        mSelectedPath = Left$(Buffer, lstrlenW(pBuffer))
        ShowDialog = DialogResult_OK
        'Free the IDList Memory !
        Call CoTaskMemFree(IDList)
    Else
        ShowDialog = DialogResult_Cancel
    End If
    Exit Function
Catch:
    If Err.Number > 0 Then MsgBox Err.Description
    ShowDialog = DialogResult_Abort
End Function

Private Function FncPtr(p As Long) As Long
    FncPtr = p
End Function

' Callbackfunction ->>>
Private Sub ICallBack_CallBack(ByVal hhwnd As Long, ByVal msg As Long, ByVal lParam As Long)
    Dim rv     As Long
    Dim lflag  As Long
    Dim Buffer As String
    Select Case msg
    Case BFFM_INITIALIZED
        If (Len(mSelectedPath) > 0) Then
            Dim SelPath As String: SelPath = mSelectedPath
            Dim pSelPath As Long: pSelPath = StrPtr(SelPath)
            rv = SendMessageW(hhwnd, BFFM_SETSELECTIONW, 1&, ByVal pSelPath)
        End If
    Case BFFM_SELCHANGED
        If (lParam <> 0&) Then
            Buffer = String$(1024, vbNullChar)
            lflag = SHGetPathFromIDList(lParam, Buffer)
            If lflag = 1 Then
                Call SendMessageW(hhwnd, BFFM_ENABLEOK, 0, ByVal 1)
            ElseIf lflag = 0 Then
                Call SendMessageW(hhwnd, BFFM_ENABLEOK, 0, ByVal 0)
            End If
            Call CoTaskMemFree(lParam)
        End If
    End Select
End Sub


