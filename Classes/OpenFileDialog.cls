VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Public NotInheritable Class OpenFileDialog
'          Inherits System.Windows.Forms.FileDialog
Option Explicit '490 Zeilen
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetOpenFileNamePreview Lib "msvfw32.dll" (ByRef lpofn As OPENFILENAME) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As WinRect) As Long
'Private Const WS_CHILD As Long = &H40000000
'Private Const WS_VISIBLE As Long = &H10000000

'Private Type WinRect
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

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

'Diese Klasse funktioniert unter VB6 genauso wie die
'Klasse OpenFileDialog von .NET
'also die wichtigsten Sachen
'
'CommonDialog
'Microsoft Common Dialog Control 6.0 (SP6)
'in: \System32\COMDLG32.OCX
'ist nur eine dünne Kapselung der dll "comdlg32.dll"
'deshalb lieber gleich die dll-verwenden, hat den Vorteil, daß comdlg32.ocx
'nicht mit dem Programm ausgeliefert werden muß,
'da die comdlg32.dll Bestandteil des Betriebssystems ist
'Private Base As New FileDialog
Private mAddExtension As Boolean
Attribute mAddExtension.VB_VarHelpID = -1
'Private mCheckFileExists As Boolean
'Private mCheckPathExists As Boolean
Private mDefaultExt As String
Attribute mDefaultExt.VB_VarHelpID = -1
'Private mDereferenceLinks As Boolean
Private mFileName As String
Attribute mFileName.VB_VarHelpID = -1
Private mFileNames As Collection    'ReadOnly nur Get
Private mFilter As String
Attribute mFilter.VB_VarHelpID = -1
Private mFilterIndex As Long
Attribute mFilterIndex.VB_VarHelpID = -1
Private mInitialDirectory As String
'Private mShowReadOnly As Boolean
Private mTitle As String
Attribute mTitle.VB_VarHelpID = -1
Private mValidateNames As Boolean
Attribute mValidateNames.VB_VarHelpID = -1
Private mFlags As Long
Private mTag As Variant

Private Sub Class_Initialize()
    Call Reset
End Sub
Public Sub Reset()
    'Call InitFlags
    mFlags = mFlags Or OFN_EXPLORER
    mFlags = mFlags Or OFN_FILEMUSTEXIST
    mFlags = mFlags Or OFN_HIDEREADONLY
    mFlags = mFlags Or OFN_PATHMUSTEXIST
    
    mAddExtension = True
    'Set mFileNames = New Collection  'ReadOnly nur Get
    Call FNsColInitClear
    mFilterIndex = 1           'nFilterIndex
    'mTitle = "Open"                '
    mValidateNames = True      'OFN_NOVALIDATE
End Sub
Private Sub FNsColInitClear()
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
Friend Property Let flags(FlgVal As Long)
    mFlags = FlgVal
End Property

'Ruft einen Wert ab, oder legt diesen fest, der angibt,
'ob einem Dateinamen im Dialogfeld automatisch eine Erweiterung
'hinzugefügt wird wenn der Benutzer keine Erweiterung angibt
Public Property Get AddExtension() As Boolean
    AddExtension = mAddExtension
End Property
Public Property Let AddExtension(BolVal As Boolean)
    mAddExtension = BolVal
End Property

'Public Overrides Property CheckFileExists() As Boolean
Public Property Get CheckFileExists() As Boolean       'OFN_FILEMUSTEXIST
    'CheckFileExists = mCheckFileExists
    CheckFileExists = mFlags And OFN_FILEMUSTEXIST
End Property
Public Property Let CheckFileExists(BolVal As Boolean)       'OFN_FILEMUSTEXIST
    'mCheckFileExists = BolVal
    mFlags = mFlags Or OFN_FILEMUSTEXIST
    If Not BolVal Then
        mFlags = mFlags Xor OFN_FILEMUSTEXIST
    End If
End Property

Public Property Get CheckPathExists() As Boolean         'OFN_PATHMUSTEXIST
     'CheckPathExists = mCheckPathExists
    CheckPathExists = mFlags And OFN_PATHMUSTEXIST
End Property
Public Property Let CheckPathExists(BolVal As Boolean)         'OFN_PATHMUSTEXIST
    'mCheckPathExists = BolVal
    mFlags = mFlags Or OFN_PATHMUSTEXIST
    If Not BolVal Then
        mFlags = mFlags Xor OFN_PATHMUSTEXIST
    End If
End Property

'Public Property DefaultExt() As String
Public Property Get DefaultExt() As String
    DefaultExt = mDefaultExt
End Property
Public Property Let DefaultExt(StrExt As String)
    mDefaultExt = StrExt
End Property

'Public Property DereferenceLinks() As Boolean
Public Property Get DereferenceLinks() As Boolean       'OFN_NODEREFERENCELINKS
    'DereferenceLinks = mDereferenceLinks
    DereferenceLinks = mFlags And OFN_NODEREFERENCELINKS
End Property
Public Property Let DereferenceLinks(BolVal As Boolean) 'OFN_NODEREFERENCELINKS
    mFlags = mFlags Or OFN_NODEREFERENCELINKS
    If BolVal Then
        mFlags = mFlags Xor OFN_NODEREFERENCELINKS
    End If
End Property

'Public Property Let DereferenceLinks(BolVal As Boolean) 'OFN_NODEREFERENCELINKS
'  mFlags = mFlags Or OFN_NODEREFERENCELINKS
'  If BolVal Then
'    mFlags = mFlags Xor OFN_NODEREFERENCELINKS
'  End If
'End Property

'Public Property FileName() As String
Public Property Get FileName() As String
    FileName = mFileName
End Property
Public Property Let FileName(StrFN As String)
    mFileName = StrFN
End Property

'Public ReadOnly Property FileNames() As String()
Public Property Get FileNames() As Collection 'ReadOnly nur Get
    Set FileNames = mFileNames
End Property

Private Sub GetFNsinCol(FNs As String)
Dim i As Integer
Dim aa As String, bb As String, Path As String
Dim PathFirst As Boolean
Dim FirstFile As Boolean
Dim Pos As Long, X As Long, xOld As Long
Dim FN As String
  PathFirst = False
  FirstFile = True
  Call FNsColInitClear
        
  'Schauen ob evt. nur eine Datei ausgewählt wurde
  Call TrimVBNullchar(FNs)
  If InStr(FNs, vbNullChar) <> 0 Then
    aa = FNs & vbNullChar
    Pos = 1
    
    Do
      'Chr$(0) ist Trennzeichen
      X = InStr(Pos, aa, vbNullChar)
      If X = xOld + 1 Then 'zwei vbnullchar hintereinander, dann aufhören
        Exit Do
      End If
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
Private Sub TrimVBNullchar(ByRef StrVal As String) 'As String
Dim i As Long
Dim sepStr As String
  For i = 1 To Len(StrVal)
    sepStr = vbNullChar + vbNullChar
    If Mid(StrVal, i, 2) = sepStr Then
      StrVal = VBA.Left$(StrVal, i - 1)
      Exit For
    End If
  Next
End Sub

'Public Property Filter() As String
Public Property Get Filter() As String
Dim i As Integer
Dim StrFlt As String
  StrFlt = mFilter
  For i = 1 To Len(StrFlt)
    If Mid(StrFlt, i, 1) = vbNullChar Then Mid(StrFlt, i, 1) = "|"
  Next
  Filter = StrFlt
End Property
Public Property Let Filter(StrVal As String)
Dim i As Integer
Dim StrFlt As String
  StrFlt = StrVal
  For i = 1 To Len(StrFlt)
    If Mid(StrFlt, i, 1) = "|" Then Mid(StrFlt, i, 1) = vbNullChar
  Next
  If Not Right(StrFlt, 2) = vbNullChar + vbNullChar Then StrFlt = StrFlt + vbNullChar + vbNullChar
  mFilter = StrFlt
End Property

'Public Property FilterIndex() As Integer
Public Property Get FilterIndex() As Long
  FilterIndex = mFilterIndex
End Property
Public Property Let FilterIndex(IntVal As Long)
  mFilterIndex = IntVal
End Property

'Ruft das Ausgangsverzeichnis ab, oder legt dieses fest
Public Property Get InitialDirectory() As String       'lpstrInitialDir
  InitialDirectory = mInitialDirectory
End Property
Public Property Let InitialDirectory(StrVal As String) 'lpstrInitialDir
  mInitialDirectory = StrVal
End Property

Public Property Get MultiSelect() As Boolean         'OFN_ALLOWMULTISELECT
  MultiSelect = mFlags And OFN_ALLOWMULTISELECT
End Property
Public Property Let MultiSelect(BolVal As Boolean)   'OFN_ALLOWMULTISELECT
  mFlags = mFlags Or OFN_ALLOWMULTISELECT
  If Not BolVal Then
    mFlags = mFlags Xor OFN_ALLOWMULTISELECT
  End If
End Property

Public Property Get ReadOnlyChecked() As Boolean       'OFN_READONLY
  ReadOnlyChecked = mFlags And OFN_READONLY
End Property
Public Property Let ReadOnlyChecked(BolVal As Boolean) 'OFN_READONLY
  mFlags = mFlags Or OFN_READONLY
  If Not BolVal Then
    mFlags = mFlags Xor OFN_READONLY
  End If
End Property

'ruft einen Wert ab, oder legt diesen fest, der angibt,
'ob das Dialogfeld das vor dem Schließen aktuelle Verzeichnis wiederherstellt,
Public Property Get RestoreDirectory() As Boolean       'OFN_NOCHANGEDIR
  RestoreDirectory = mFlags And OFN_NOCHANGEDIR
End Property
Public Property Let RestoreDirectory(BolVal As Boolean) 'OFN_NOCHANGEDIR
  mFlags = mFlags Or OFN_NOCHANGEDIR
  If Not BolVal Then
    mFlags = mFlags Xor OFN_NOCHANGEDIR
  End If
End Property

'Public Property ShowHelp() As Boolean
Public Property Get ShowHelp() As Boolean           'OFN_SHOWHELP
  ShowHelp = mFlags And OFN_SHOWHELP
End Property
Public Property Let ShowHelp(BolVal As Boolean)     'OFN_SHOWHELP
  mFlags = mFlags Or OFN_SHOWHELP
  If Not BolVal Then
    mFlags = mFlags Xor OFN_SHOWHELP
  End If
End Property

'Public Property ShowReadOnly() As Boolean
Public Property Get ShowReadOnly() As Boolean       'OFN_HIDEREADONLY
  'ShowReadOnly = mShowReadOnly
  ShowReadOnly = mFlags And OFN_HIDEREADONLY
End Property
Public Property Let ShowReadOnly(BolVal As Boolean) 'OFN_HIDEREADONLY
  'mShowReadOnly = BolVal
  mFlags = mFlags Or OFN_HIDEREADONLY
  If BolVal Then
    mFlags = mFlags Xor OFN_HIDEREADONLY
  End If
End Property
'Private Sub mShowReadOnly_OnLet(ByVal sender As Object)
'  mFlags = mFlags Or OFN_HIDEREADONLY
'  If BolVal Then
'    mFlags = mFlags Xor OFN_HIDEREADONLY
'  End If
'End Sub

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
  ToString = "MBO.Windows.Forms.OpenFileDialog: " & "Title: " & mTitle & " , FileName: " & mFileName
End Function

Public Property Get Title() As String       'lpstrTitle
  Title = mTitle
End Property
Public Property Let Title(StrVal As String) 'lpstrTitle
  mTitle = StrVal
End Property

'ob das Dialogfeld nur gültige Win32-Dateinamen akzeptiert
Public Property Get ValidateNames() As Boolean       'OFN_NOVALIDATE
  ValidateNames = mFlags And OFN_NOVALIDATE
End Property
Public Property Let ValidateNames(BolVal As Boolean) 'OFN_NOVALIDATE
  mFlags = mFlags Or OFN_NOVALIDATE
  If BolVal Then
    mFlags = mFlags Xor OFN_NOVALIDATE
  End If
End Property

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult

Try: On Error GoTo Catch
  
    Dim myOwnerHwnd As Long
    If Not IsMissing(Frm) Then
        myOwnerHwnd = Frm.hwnd
    Else
        myOwnerHwnd = GetActiveWindow
        'myOwnerHwnd = GetCenteredHwnd 'see below 'scheiße das bringt nix!!!
        'myOwnerHwnd = GetDesktopWindow 'GetActiveWindow
    End If
    Dim myOwnerhInst As Long: myOwnerhInst = App.hInstance
    
    Dim mMaxLen As Long:  mMaxLen = 32767 - Len(mFileName)
    Dim Buffer  As String: Buffer = mFileName & String$(mMaxLen, vbNullChar)
    
    Dim OFN As OPENFILENAME
    With OFN
        .lStructSize = LenB(OFN)
        .hwndOwner = myOwnerHwnd
        .hInstance = myOwnerhInst
        .lpstrFilter = mFilter
        '.lpstrCustomFilter = mDefaultExt
        'nMaxCustFilter
        .nFilterIndex = mFilterIndex '(FilterIndex)
        .lpstrFile = Buffer
        .nMaxFile = Len(Buffer) '32767
        'lpstrFileTitle 'der NurDateiName ohne Extension
        'nMaxFileTitle
        .lpstrInitialDir = mInitialDirectory
        .lpstrTitle = mTitle
        .flags = mFlags
        'nFileOffset
        'nFileExtension
        .lpstrDefExt = mDefaultExt
        'lCustData
        'lpfnHook
        'lpTemplateName
    End With
    Dim Result As Long: Result = GetOpenFileName(OFN)
    'Result = GetOpenFileNamePreview(OFN)
    If Result <> 0 Then
        'man muß am String erkennen, ob nur eine oder mehrere Dateien ausgewählt wurden
        Call GetFNsinCol(OFN.lpstrFile)
        ShowDialog = vbOK
    Else
        If mFlags And OFN_ALLOWMULTISELECT Then
            Call GetFNsinCol(OFN.lpstrFile)
        End If
        ShowDialog = vbCancel
        'oder zuviele Dateien
    End If
Finally:
    'Call DestroyWindow(myOwnerHwnd)
    'Debug.Print "DestroyWindow " & myOwnerHwnd
    On Error GoTo Catch
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

