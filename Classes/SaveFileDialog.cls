VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaveFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public MustInherit Class FileDialog
'          Inherits System.Windows.Forms.CommonDialog
'     Member von: System.Windows.Forms
Option Explicit '521 Zeilen   '2021-08-07: 613 Zeilen
'Diese Klasse funktioniert unter VB6 genauso wie die
'Klasse FileDialog von .NET
'also bis jetzt die wichtigsten Sachen
'CommonDialog
'Microsoft Common Dialog Control 6.0 (SP6)
'in: \System32\COMDLG32.OCX
'ist nur eine dünne Kapselung der dll "comdlg32.dll"
'deshalb lieber gleich die dll-verwenden, hat den Vorteil, daß comdlg32.ocx
'nicht mit dem Programm ausgeliefert werden muß,
'da die comdlg32.dll Bestandteil des Betriebssystems ist

Private Const OFN_SHAREWARN             As Long = 0&
Private Const OFN_SHARENOWARN           As Long = 1&
Private Const OFN_SHAREFALLTHROUGH      As Long = 2&

Private Const OFN_EX_NOPLACESBAR        As Long = &H1

Private Const OFN_READONLY              As Long = &H1&
Private Const OFN_OVERWRITEPROMPT       As Long = &H2&
Private Const OFN_HIDEREADONLY          As Long = &H4&
Private Const OFN_NOCHANGEDIR           As Long = &H8&

Private Const OFN_SHOWHELP              As Long = &H10&
Private Const OFN_ENABLEHOOK            As Long = &H20&
Private Const OFN_ENABLETEMPLATE        As Long = &H40&
Private Const OFN_ENABLETEMPLATEHANDLE  As Long = &H80&

Private Const OFN_NOVALIDATE            As Long = &H100&
Private Const OFN_ALLOWMULTISELECT      As Long = &H200&
Private Const OFN_EXTENSIONDIFFERENT    As Long = &H400&
Private Const OFN_PATHMUSTEXIST         As Long = &H800&

Private Const OFN_FILEMUSTEXIST         As Long = &H1000&
Private Const OFN_CREATEPROMPT          As Long = &H2000&
Private Const OFN_SHAREAWARE            As Long = &H4000&
Private Const OFN_NOREADONLYRETURN      As Long = &H8000&

Private Const OFN_NOTESTFILECREATE      As Long = &H10000
Private Const OFN_NONETWORKBUTTON       As Long = &H20000
Private Const OFN_NOLONGNAMES           As Long = &H40000
Private Const OFN_EXPLORER              As Long = &H80000

Private Const OFN_NODEREFERENCELINKS    As Long = &H100000
Private Const OFN_LONGNAMES             As Long = &H200000
Private Const OFN_ENABLEINCLUDENOTIFY   As Long = &H400000
Private Const OFN_ENABLESIZING          As Long = &H800000

Private Const OFN_USEMONIKERS           As Long = &H1000000
Private Const OFN_DONTADDTORECENT       As Long = &H2000000
Private Const OFN_FORCESHOWHIDDEN       As Long = &H10000000

'Private Const WS_CHILD   As Long = &H40000000
'Private Const WS_VISIBLE As Long = &H10000000
'Public Enum DialogResult 'auch in OpenFileDialog hier aktivieren wenn nur SaveFileDialog gebraucht wird
'    DialogResultNone = 0
'    DialogResultOK = 1
'    DialogResultCancel = 2
'    DialogResultAbort = 3
'    DialogResultRetry = 4
'    DialogResultIgnore = 5
'    DialogResultYes = 6
'    DialogResultNo = 7
'End Enum
'Private Type WinRect
'  Left   As Long
'  Top    As Long
'  Right  As Long
'  Bottom As Long
'End Type
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
Private Type OPENFILENAMEW                 ' x86    ' x64
    lStructSize       As Long              '   4    '   4 + 4Padb
    hwndOwner         As LongPtr ' Long    '   4    '   8
    hInstance         As LongPtr ' Long    '   4    '   8
    lpstrFilter       As LongPtr ' String  '   4    '   8
    lpstrCustomFilter As LongPtr ' String  '   4    '   8
    nMaxCustFilter    As Long              '   4    '   4
    nFilterIndex      As Long              '   4    '   4
    lpstrFile         As LongPtr ' String  '   4    '   8
    nMaxFile          As Long              '   4    '   4 + 4Padb
    lpstrFileTitle    As LongPtr ' String  '   4    '   8
    nMaxFileTitle     As Long              '   4    '   4 + 4Padb
    lpstrInitialDir   As LongPtr ' String  '   4    '   8
    lpstrTitle        As LongPtr ' String  '   4    '   8
    flags             As Long              '   4    '   4
    nFileOffset       As Integer           '   2    '   2
    nFileExtension    As Integer           '   2    '   2
    lpstrDefExt       As LongPtr ' String  '   4    '   8
    lCustData         As LongPtr ' Long    '   4    '   8
    lpfnHook          As LongPtr ' Long    '   4    '   8
    lpTemplateName    As LongPtr ' String  '   4    '   8
End Type                               ' Sum: 76    ' 136

#If VBA7 Then
Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32" (ByRef pOpenfilename As OPENFILENAMEW) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
'Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
'Private Declare Function GetOpenFileNamePreview Lib "msvfw32" (ByRef lpofn As OPENFILENAME) As Long
'Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As Long
'Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As WinRect) As Long
#Else
Private Declare Function GetSaveFileNameW Lib "comdlg32" (ByRef pOpenfilename As OPENFILENAMEW) As Long
Private Declare Function GetActiveWindow Lib "user32" () As LongPtr
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetOpenFileNamePreview Lib "msvfw32" (ByRef lpofn As OPENFILENAME) As Long
'Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As WinRect) As Long
#End If

'Private Base As New FileDialog
Private mAddExtension     As Boolean
'Private mCheckFileExists  As Boolean
'Private mCheckPathExists  As Boolean
Private mDefaultExt       As String
'Private mDereferenceLinks As Boolean
Private mFileName         As String
Private mFileNames        As Collection ' List Of FileName   'ReadOnly nur Get
Private mFilter           As String
Attribute mFilter.VB_VarHelpID = -1
Private mFilterIndex      As Long
Attribute mFilterIndex.VB_VarHelpID = -1
Private mInitialDirectory As String
'Private mReadOnlyChecked  As Boolean
'Private mRestoreDirectory As Boolean
Private mShowHelp         As Boolean
Attribute mShowHelp.VB_VarHelpID = -1
Private mShowReadOnly     As Boolean
Attribute mShowReadOnly.VB_VarHelpID = -1
Private mTitle            As String
Private mValidateNames    As Boolean
Attribute mValidateNames.VB_VarHelpID = -1
Private mFlags            As Long
Private mTag              As Variant

'Public Event FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
'Public Event FileOk(ByVal sender As Object, ByVal e As System.CancelEventArgs)

Private Sub Class_Initialize()
    Call Reset
End Sub
'Protected Sub OnFileOk(ByVal e As System.ComponentModel.CancelEventArgs)

'Public Overrides Sub Reset()
Public Sub Reset()
    'Call InitFlags
    mFlags = mFlags Or OFN_EXPLORER
    mFlags = mFlags Or OFN_FILEMUSTEXIST
    mFlags = mFlags Or OFN_HIDEREADONLY
    mFlags = mFlags Or OFN_PATHMUSTEXIST
    
    mAddExtension = True
    FileNames_Clear
    mFilterIndex = 1           'nFilterIndex
    mValidateNames = True      'OFN_NOVALIDATE
End Sub
Private Sub FileNames_Clear()
    Set mFileNames = New Collection
End Sub

Private Sub InitFlags() 'in Modul ModCommonDialogs.bas
    'mFlags = mFlags Or OFN_ALLOWMULTISELECT
    'mFlags = mFlags Or OFN_CREATEPROMPT
    'mFlags = mFlags Or OFN_DONTADDTORECENT
    'mFlags = mFlags Or OFN_ENABLEHOOK
    'mFlags = mFlags Or OFN_ENABLEINCLUDENOTIFY
    'mFlags = mFlags Or OFN_ENABLESIZING
    'mFlags = mFlags Or OFN_ENABLETEMPLATE
    'mFlags = mFlags Or OFN_ENABLETEMPLATEHANDLE
    'mFlags = mFlags Or OFN_EX_NOPLACESBAR
    mFlags = mFlags Or OFN_EXPLORER
    'mFlags = mFlags Or OFN_EXTENSIONDIFFERENT
    mFlags = mFlags Or OFN_FILEMUSTEXIST
    'mFlags = mFlags Or OFN_FORCESHOWHIDDEN
    mFlags = mFlags Or OFN_HIDEREADONLY
    'mFlags = mFlags Or OFN_LONGNAMES
    'mFlags = mFlags Or OFN_NOCHANGEDIR
    'mFlags = mFlags Or OFN_NODEREFERENCELINKS
    'mFlags = mFlags Or OFN_NOLONGNAMES
    'mFlags = mFlags Or OFN_NONETWORKBUTTON
    'mFlags = mFlags Or OFN_NOREADONLYRETURN
    'mFlags = mFlags Or OFN_NOTESTFILECREATE
    'mFlags = mFlags Or OFN_NOVALIDATE
    'mFlags = mFlags Or OFN_OVERWRITEPROMPT
    mFlags = mFlags Or OFN_PATHMUSTEXIST
    'mFlags = mFlags Or OFN_READONLY
    'mFlags = mFlags Or OFN_SHAREAWARE
    'mFlags = mFlags Or OFN_SHAREFALLTHROUGH
    'mFlags = mFlags Or OFN_SHARENOWARN
    'mFlags = mFlags Or OFN_SHAREWARN
    'mFlags = mFlags Or OFN_SHOWHELP
    'mFlags = mFlags Or OFN_USEMONIKERS
End Sub

Friend Property Get flags() As Long
    flags = mFlags
End Property
Friend Property Let flags(ByVal Value As Long)
    mFlags = Value
End Property

'Protected Overrides Function HookProc(ByVal hWnd As System.IntPtr, ByVal msg As Integer, ByVal wparam As System.IntPtr, ByVal lparam As System.IntPtr) As System.IntPtr
'Protected Overrides Function RunDialog(ByVal hWndOwner As System.IntPtr) As Boolean
'Public Function RunDialog(ByVal hwndOwner As Long) As Boolean
'  'Set RunDialog = New Boolean
'End Function
'Public Function ToString() As String
'    'Set ToString = New String
'    'was soll hier ausgegeben werden?
'End Function

'Ruft einen Wert ab,  oder legt diesen fest, der angibt,
'ob einem Dateinamen im Dialogfeld automatisch eine Erweiterung
'hinzugefügt wird wenn der Benutzer keine Erweiterung angibt
Public Property Get AddExtension() As Boolean
    AddExtension = mAddExtension
End Property
Public Property Let AddExtension(ByVal Value As Boolean)
    mAddExtension = Value
End Property

'Public Overridable Property CheckFileExists() As Boolean
Public Property Get CheckFileExists() As Boolean               'OFN_FILEMUSTEXIST
    CheckFileExists = mFlags And OFN_FILEMUSTEXIST
End Property
Public Property Let CheckFileExists(ByVal Value As Boolean)    'OFN_FILEMUSTEXIST
    mFlags = mFlags Or OFN_FILEMUSTEXIST
    If Not Value Then mFlags = mFlags Xor OFN_FILEMUSTEXIST
End Property

Public Property Get CheckPathExists() As Boolean               'OFN_PATHMUSTEXIST
    CheckPathExists = mFlags And OFN_PATHMUSTEXIST
End Property
Public Property Let CheckPathExists(ByVal Value As Boolean)    'OFN_PATHMUSTEXIST
    mFlags = mFlags Or OFN_PATHMUSTEXIST
    If Not Value Then mFlags = mFlags Xor OFN_PATHMUSTEXIST
End Property

'Public Property DefaultExt() As String
Public Property Get DefaultExt() As String
    DefaultExt = mDefaultExt
End Property
Public Property Let DefaultExt(Extension As String)
    mDefaultExt = Extension
End Property

'Public Property DereferenceLinks() As Boolean
Public Property Get DereferenceLinks() As Boolean              'OFN_NODEREFERENCELINKS
    DereferenceLinks = mFlags And OFN_NODEREFERENCELINKS
End Property
Public Property Let DereferenceLinks(ByVal Value As Boolean)   'OFN_NODEREFERENCELINKS
    mFlags = mFlags Or OFN_NODEREFERENCELINKS
    If Value Then mFlags = mFlags Xor OFN_NODEREFERENCELINKS
End Property
'Private Sub mDereferenceLinks_OnLet(ByVal sender As Object)
'    Dim Value As Boolean: b = sender '???
'    'If sender Is Boolean Then
'    mFlags = mFlags Or OFN_NODEREFERENCELINKS
'    'Achtung hier nicht Not !! weil: Not No 2*Verneinung
'    If Value Then mFlags = mFlags Xor OFN_NODEREFERENCELINKS
'End Sub

Public Property Get FileName() As String
    FileName = mFileName
End Property
Public Property Let FileName(ByVal Value As String)
    mFileName = Value
End Property

'Public ReadOnly Property FileNames() As String()
Public Property Get FileNames() As Collection 'ReadOnly nur Get
    Set FileNames = mFileNames
End Property

Private Sub FileNames_Parse(FNs As String)
    Dim aa As String, bb As String, Path As String
    Dim X As Long, xOld As Long
    Dim FN As String
    
    Dim PathFirst As Boolean: PathFirst = False
    Dim FirstFile As Boolean: FirstFile = True
    
    FileNames_Clear

    
    'Schauen ob evt. nur eine Datei ausgewählt wurde
    Call TrimVBNullchar(FNs)
    If InStr(FNs, vbNullChar) <> 0 Then
        aa = FNs & vbNullChar
        Dim Pos As Long: Pos = 1
        
        Do
            'Chr$(0) ist Trennzeichen
            X = InStr(Pos, aa, vbNullChar)
            'zwei vbnullchar hintereinander, dann aufhören
            If X = xOld + 1 Then Exit Do
            xOld = X
            If X = 0 Then Exit Do
            bb = Mid(aa, Pos, X - Pos)
               
            If Not PathFirst Then
                'Im ersten Teil befindet sich immer der Pfad
                Path = bb
                If Not VBA.Right$(Path, 1) = "\" Then Path = Path & "\"
                PathFirst = True
            Else
                'Pfad ist bereits bekannt, es folgen Dateinamen
                If FirstFile Then
                    mFileName = Path & bb
                    FirstFile = False
                End If
                Call mFileNames.Add(Path & bb)
            End If
            Pos = X + 1
        Loop
    Else
        'Es wurde nur eine Datei ausgewählt
        'Call mFileNames.Add(FNs)
        FN = FNs
        mFileName = FN
    End If
End Sub


Private Sub TrimVBNullchar(ByRef Value_out As String) 'As String
    Dim sSep As String: sSep = vbNullChar + vbNullChar
    Dim i As Long
    For i = 1 To Len(Value_out)
        If Mid(Value_out, i, 2) = sSep Then
            Value_out = VBA.Left$(Value_out, i - 1)
            Exit For
        End If
    Next
End Sub

Public Property Get Filter() As String
    Filter = Replace(mFilter, vbNullChar, "|")
'    Dim i As Long
'    Dim StrFlt As String
'    StrFlt = mFilter
'    For i = 1 To Len(StrFlt)
'        If Mid(StrFlt, i, 1) = vbNullChar Then Mid(StrFlt, i, 1) = "|"
'    Next
'    Filter = StrFlt
End Property
Public Property Let Filter(ByVal Value As String)
    'Debug.Print Value
    mFilter = Replace(Value, "|", vbNullChar)
    'Debug.Print mFilter
'Dim i As Integer
'Dim StrFlt As String
'  StrFlt = StrVal
'  Debug.Print StrFlt
'  For i = 1 To Len(StrFlt)
'    If Mid(StrFlt, i, 1) = "|" Then Mid(StrFlt, i, 1) = vbNullChar
'  Next
'  If Not Right(StrFlt, 2) = vbNullChar + vbNullChar Then StrFlt = StrFlt + vbNullChar + vbNullChar
'  mFilter = StrFlt
'  Debug.Print mFilter
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = mFilterIndex
End Property
Public Property Let FilterIndex(ByVal Value As Long)
    mFilterIndex = Value
End Property

'Ruft das Ausgangsverzeichnis ab, oder legt dieses fest
Public Property Get InitialDirectory() As String       'lpstrInitialDir
    InitialDirectory = mInitialDirectory
End Property
Public Property Let InitialDirectory(ByVal Value As String) 'lpstrInitialDir
    mInitialDirectory = Value
End Property

'MultiSelect hat nur der Opendialog nicht im Savedialog möglich
'ReadOnlyChecked hat nur der Opendialog nicht im Savedialog möglich
'Public Property Get MultiSelect() As Boolean                'OFN_ALLOWMULTISELECT
'    MultiSelect = mFlags And OFN_ALLOWMULTISELECT
'End Property
'Public Property Let MultiSelect(ByVal Value As Boolean)     'OFN_ALLOWMULTISELECT
'    mFlags = mFlags Or OFN_ALLOWMULTISELECT
'    If Not Value Then mFlags = mFlags Xor OFN_ALLOWMULTISELECT
'End Property
'Public Property Get ReadOnlyChecked() As Boolean            'OFN_READONLY
'    ReadOnlyChecked = mFlags And OFN_READONLY
'End Property
'Public Property Let ReadOnlyChecked(ByVal Value As Boolean) 'OFN_READONLY
'    mFlags = mFlags Or OFN_READONLY
'    If Not Value Then mFlags = mFlags Xor OFN_READONLY
'End Property

'ruft einen Wert ab, oder legt diesen fest, der angibt,
'ob das Dialogfeld vor dem Schließen das aktuelle Verzeichnis wiederherstellt,
Public Property Get RestoreDirectory() As Boolean            'OFN_NOCHANGEDIR
    RestoreDirectory = mFlags And OFN_NOCHANGEDIR
End Property
Public Property Let RestoreDirectory(ByVal Value As Boolean) 'OFN_NOCHANGEDIR
    mFlags = mFlags Or OFN_NOCHANGEDIR
    If Not Value Then mFlags = mFlags Xor OFN_NOCHANGEDIR
End Property

Public Property Get ShowHelp() As Boolean                    'OFN_SHOWHELP
    ShowHelp = mFlags And OFN_SHOWHELP
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)         'OFN_SHOWHELP
    mFlags = mFlags Or OFN_SHOWHELP
    If Not Value Then mFlags = mFlags Xor OFN_SHOWHELP
End Property

'ShowReadOnly hat nur der Opendialog nicht im Savedialog möglich
'Public Property Get ShowReadOnly() As Boolean       'OFN_HIDEREADONLY
'    ShowReadOnly = mFlags And OFN_HIDEREADONLY
'End Property
'Public Property Let ShowReadOnly(ByVal Value As Boolean) 'OFN_HIDEREADONLY
'    mFlags = mFlags Or OFN_HIDEREADONLY
'    If Value Then mFlags = mFlags Xor OFN_HIDEREADONLY
'End Property

'wie hinlänglich bekannt, füg einfach irgendwas hinzu
'wird intern nicht verwendet, kann extern verwendet werden
Public Property Get Tag() As Variant
    Tag = mTag
End Property
Public Property Let Tag(Value As Variant)
    mTag = Value
End Property

Public Function ToString() As String
    ToString = "MBO.Windows.Forms.SaveFileDialog: " & "Title: " & mTitle & " , FileName: " & mFileName
End Function

Public Property Get Title() As String       'lpstrTitle
    Title = mTitle
End Property
Public Property Let Title(StrVal As String) 'lpstrTitle
    mTitle = StrVal
End Property

'ob das Dialogfeld nur gültige Win32-Dateinamen akzeptiert
Public Property Get ValidateNames() As Boolean            'OFN_NOVALIDATE
    ValidateNames = mFlags And OFN_NOVALIDATE
End Property
Public Property Let ValidateNames(ByVal Value As Boolean) 'OFN_NOVALIDATE
    mFlags = mFlags Or OFN_NOVALIDATE
    If Value Then mFlags = mFlags Xor OFN_NOVALIDATE
End Property

Private Function GethWnd(Frm As Variant) As LongPtr
Try: On Error GoTo Catch
    GethWnd = Frm.hwnd
    Exit Function
Catch:
    GethWnd = GetActiveWindow
End Function

Private Function GethInst() As LongPtr
#If VBA6 Or VBA7 Then
    GethInst = App.hInstance
#Else
    GethInst = Application.hInstPtr
#End If
End Function

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult

Try: On Error GoTo Catch
  
    Dim aOwnerhWnd  As LongPtr: aOwnerhWnd = GethWnd(Frm)
    Dim aOwnerhInst As LongPtr: aOwnerhInst = GethInst
    
    Dim mMaxLen As Long:   mMaxLen = 32767 - Len(mFileName)
    Dim Buffer  As String:  Buffer = mFileName & String$(mMaxLen, vbNullChar)
    Dim aFilter As String: aFilter = mFilter & vbNullChar & vbNullChar
    Dim OFN As OPENFILENAMEW
    With OFN
        .lStructSize = LenB(OFN)
        .hwndOwner = aOwnerhWnd
        .hInstance = aOwnerhInst
        .lpstrFilter = StrPtr(aFilter)
        '.lpstrCustomFilter = mDefaultExt
        '.nMaxCustFilter
        .nFilterIndex = CLng(FilterIndex)
        .lpstrFile = StrPtr(Buffer)
        .nMaxFile = Len(Buffer) '32767
        '.lpstrFileTitle
        '.nMaxFileTitle
        .lpstrInitialDir = StrPtr(mInitialDirectory)
        .lpstrTitle = StrPtr(Title)
        .flags = mFlags
        'nFileOffset
        'nFileExtension = 1
        If mAddExtension Then
            If Len(mDefaultExt) = 0 Then
              '
            End If
            .lpstrDefExt = StrPtr(mDefaultExt)
        End If
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
    End With
    Dim hr As Long: hr = GetSaveFileNameW(OFN) '<<<<<<------ here Get_*_SAVE_*_FileNameW
    'hr = GetSaveFileNamePreview(OFN)
    If hr Then
        'man muß am String erkennen, ob nur eine oder mehrere Dateien ausgewählt wurden
        FileNames_Parse Buffer 'OFN.lpstrFile)
        ShowDialog = vbOK
    Else
        'If mFlags And OFN_ALLOWMULTISELECT Then
            FileNames_Parse Buffer
        'End If
        ShowDialog = vbCancel
        'oder zuviele Dateien
    End If
    Exit Function
Catch:
    If Err.Number = 32755 Then
        ShowDialog = vbCancel
    Else
        ShowDialog = vbAbort
        MsgBox CStr(Err.Number) + " " + Err.Description
    End If
End Function

'Private Function GetCenteredHwnd() As Long
'Dim mLeft As Long, mTop As Long, mWidth  As Long, mHeight As Long
'Dim mDesktopHWnd As Long:  mDesktopHWnd = GetActiveWindow 'GetDesktopWindow
'Dim mTitle  As String:   mTitle = vbNullString
'Dim mWStyle As Long:     mWStyle = GetWStyle
'Dim mWExStyle As Long:   mWExStyle = GetWExStyle
'Dim mStrClass As String: mStrClass = "STATIC" '"MDICLIENT" '
'  mWidth = 563&  'minimale Breite des OpenFileDialogs unter WinXP
'  mHeight = 419& 'minimale Hoehe  des OpenFileDialogs unter WinXP
'  mLeft = CLng((Screen.Width / Screen.TwipsPerPixelX) / 2 - mWidth / 2 - 3) '1024:=> 230.5
'  mTop = CLng((Screen.Height / Screen.TwipsPerPixelY) / 2 - mHeight / 2 - 18) ' 768:=> 174.5
'  GetCenteredHwnd = CreateWindowExA(mWExStyle, mStrClass, mTitle, mWStyle, mLeft, mTop, mWidth, mHeight, mDesktopHWnd, 0&, App.hInstance, 0&)
'  Debug.Print GetCenteredHwnd
'End Function

'Private Function GetWStyle() As WindowStyle
  'GetWStyle = GetWStyle Or WS_CAPTION
  'GetWStyle = GetWStyle Or WS_ACTIVECAPTION
  'GetWStyle = GetWStyle Or WS_SYSMENU
  'GetWStyle = GetWStyle Or WS_MAXIMIZEBOX
  'GetWStyle = GetWStyle Or WS_MINIMIZEBOX
  
  'GetWStyle = GetWStyle Or WS_HSCROLL
  'GetWStyle = GetWStyle Or WS_VSCROLL
  
  'GetWStyle = GetWStyle Or WS_TABSTOP
  'GetWStyle = GetWStyle Or WS_GROUP
  
  'GetWStyle = GetWStyle Or WS_DLGFRAME
  'GetWStyle = GetWStyle Or WS_THICKFRAME
  
  'GetWStyle = GetWStyle Or WS_OVERLAPPED
  'GetWStyle = GetWStyle Or WS_BORDER
  'GetWStyle = GetWStyle Or WS_MAXIMIZE
  'GetWStyle = GetWStyle Or WS_MINIMIZE
  'GetWStyle = GetWStyle Or WS_CLIPCHILDREN
  'GetWStyle = GetWStyle Or WS_CLIPSIBLINGS
  'GetWStyle = GetWStyle Or WS_DISABLED
  'GetWStyle = GetWStyle Or WS_VISIBLE
  'GetWStyle = GetWStyle Or WS_CHILD
  'GetWStyle = GetWStyle Or WS_POPUP
'End Function
'Private Function GetWExStyle() As WindowExStyle
  'GetWExStyle = GetWExStyle Or WS_EX_LEFTSCROLLBAR
  'GetWExStyle = GetWExStyle Or WS_EX_RIGHTSCROLLBAR
  'GetWExStyle = GetWExStyle Or WS_EX_DLGMODALFRAME
  'GetWExStyle = GetWExStyle Or WS_EX_NOPARENTNOTIFY
  'GetWExStyle = GetWExStyle Or WS_EX_TOPMOST
  'GetWExStyle = GetWExStyle Or WS_EX_ACCEPTFILES
  'GetWExStyle = GetWExStyle Or WS_EX_TRANSPARENT
  'GetWExStyle = GetWExStyle Or WS_EX_MDICHILD
  'GetWExStyle = GetWExStyle Or WS_EX_WINDOWEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_CLIENTEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_CONTEXTHELP
  'GetWExStyle = GetWExStyle Or WS_EX_RIGHT
  'GetWExStyle = GetWExStyle Or WS_EX_CONTROLPARENT
  'GetWExStyle = GetWExStyle Or WS_EX_STATICEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_APPWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_TOOLWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_LAYERED
  'GetWExStyle = GetWExStyle Or WS_EX_NOINHERITLAYOUT
  'GetWExStyle = GetWExStyle Or WS_EX_NOACTIVATE
  'GetWExStyle = GetWExStyle Or WS_EX_OVERLAPPEDWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_PALETTEWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_LEFT
  'GetWExStyle = GetWExStyle Or WS_EX_LTRREADING
  'GetWExStyle = GetWExStyle Or WS_EX_RTLREADING
  'GetWExStyle = GetWExStyle Or WS_EX_LAYOUTRTL
'End Function

'    OFD.AddExtension()
'    OFD.CheckFileExists()
'    OFD.CheckPathExists()
'    OFD.Container()
'    OFD.CreateObjRef()
'    OFD.DefaultExt()
'    OFD.DereferenceLinks()
'    OFD.Dispose()
'    OFD.FileName()
'    OFD.FileNames()
'    OFD.Filter()
'    OFD.FilterIndex()
'    OFD.GetLifetimeService()
'    OFD.GetType()
'    OFD.InitialDirectory()
'    OFD.InitializeLifetimeService()
'    OFD.Multiselect()
'    OFD.OpenFile()
'    OFD.ReadOnlyChecked()
'    OFD.Reset()
'    OFD.RestoreDirectory()
'    OFD.ShowDialog()
'    OFD.ShowHelp()
'    OFD.ShowReadOnly()
'    OFD.Site()
'    OFD.Title()
'    OFD.ToString()
'    OFD.ValidateNames()
