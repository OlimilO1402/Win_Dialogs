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
Option Explicit '521 Zeilen
'Diese Klasse funktioniert unter VB6 genauso wie die
'Klasse FileDialog von .NET
'also bis jetzt die wichtigsten Sachen
'Private WithEvents
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetOpenFileNamePreview Lib "msvfw32.dll" (ByRef lpofn As OPENFILENAME) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
'Public Enum DialogResult 'auch in OpenFileDialog hier aktivieren wenn nur SaveFileDialog gebraucht wird
'  DialogResultNone = 0
'  DialogResultOK = 1
'  DialogResultCancel = 2
'  DialogResultAbort = 3
'  DialogResultRetry = 4
'  DialogResultIgnore = 5
'  DialogResultYes = 6
'  DialogResultNo = 7
'End Enum

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
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

Private mAddExtension As Boolean
Private mCheckFileExists As Boolean
Private mCheckPathExists As Boolean
Private mDefaultExt As String
'Private WithEvents
Private mDereferenceLinks As Boolean
Private mFileName As String
Private mFileNames As Collection    'ReadOnly nur Get
Private mFilter As String
Attribute mFilter.VB_VarHelpID = -1
Private mStrFilter As String
Private mFilterIndex As Long
Attribute mFilterIndex.VB_VarHelpID = -1
'Private WithEvents
Private mInitialDirectory As String
Private mReadOnlyChecked As Boolean
Attribute mReadOnlyChecked.VB_VarHelpID = -1
Private mRestoreDirectory As Boolean
Private mShowHelp As Boolean
Attribute mShowHelp.VB_VarHelpID = -1
Private mShowReadOnly As Boolean
Attribute mShowReadOnly.VB_VarHelpID = -1
'Private WithEvents
Private mTitle As String
Private mValidateNames As Boolean
Attribute mValidateNames.VB_VarHelpID = -1
Private mFlags As Long
'Public Event FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
'Public Event FileOk(ByVal sender As Object, ByVal e As System.CancelEventArgs)

Private Sub Class_Initialize()
  Call Reset
End Sub
'Protected Sub OnFileOk(ByVal e As System.ComponentModel.CancelEventArgs)

'Public Overrides Sub Reset()
Public Sub Reset()
'  Call InitFlags
  mFlags = mFlags Or OFN_EXPLORER
  mFlags = mFlags Or OFN_FILEMUSTEXIST
  mFlags = mFlags Or OFN_HIDEREADONLY
  mFlags = mFlags Or OFN_PATHMUSTEXIST

  mAddExtension = True
  Call FNsColInitClear 'Set mFileNames = New Collection   'ReadOnly nur GetSet mFileNames = New Collection   'ReadOnly nur Get
  mFilterIndex = 1           'nFilterIndex
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

'Protected Overrides Function HookProc(ByVal hWnd As System.IntPtr, ByVal msg As Integer, ByVal wparam As System.IntPtr, ByVal lparam As System.IntPtr) As System.IntPtr

'Protected Overrides Function RunDialog(ByVal hWndOwner As System.IntPtr) As Boolean
'Public Function RunDialog(ByVal hwndOwner As Long) As Boolean
'  'Set RunDialog = New Boolean
'
'End Function
Public Function ToString() As String
  'Set ToString = New String
  'was soll hier ausgegeben werden?
End Function

'Ruft einen Wert ab,  oder legt diesen fest, der angibt,
'ob einem Dateinamen im Dialogfeld automatisch eine Erweiterung hinzugefügt wird wenn der Benutzer keine Erweiterung angibt
'Public Property AddExtension() As Boolean
Public Property Get AddExtension() As Boolean
  AddExtension = mAddExtension
End Property
Public Property Let AddExtension(BolVal As Boolean)
  mAddExtension = BolVal
End Property

'Public Overridable Property CheckFileExists() As Boolean
Public Property Get CheckFileExists() As Boolean       'OFN_FILEMUSTEXIST
  'CheckFileExists = mFlags And OFN_FILEMUSTEXIST
  CheckFileExists = mCheckFileExists
End Property
Public Property Let CheckFileExists(BolVal As Boolean) 'OFN_FILEMUSTEXIST
  mFlags = mFlags Or OFN_FILEMUSTEXIST
  If Not BolVal Then
    mFlags = mFlags Xor OFN_FILEMUSTEXIST
  End If
End Property
'Private Sub mCheckFileExists_OnLet(ByVal sender As Object)
'Dim BolVal As Boolean
'  'If sender Is Boolean Then
'  BolVal = sender
'  mFlags = mFlags Or OFN_FILEMUSTEXIST
'  If Not BolVal Then
'    mFlags = mFlags Xor OFN_FILEMUSTEXIST
'  End If
'End Sub

'Public Property CheckPathExists() As Boolean
Public Property Get CheckPathExists() As Boolean         'OFN_PATHMUSTEXIST
  'CheckPathExists = mFlags And OFN_PATHMUSTEXIST
  CheckPathExists = mFlags And OFN_PATHMUSTEXIST
End Property
Public Property Let CheckPathExists(BolVal As Boolean)   'OFN_PATHMUSTEXIST
  mFlags = mFlags Or OFN_PATHMUSTEXIST
  If Not BolVal Then
    mFlags = mFlags Xor OFN_PATHMUSTEXIST
  End If
End Property
'Private Sub mCheckPathExists_OnLet(ByVal sender As Object)
'Dim BolVal As Boolean
'  'If sender Is Boolean Then
'  BolVal = sender
'  mFlags = mFlags Or OFN_FILEMUSTEXIST
'  If Not BolVal Then
'    mFlags = mFlags Xor OFN_FILEMUSTEXIST
'  End If
'End Sub

'Public Property DefaultExt() As String
Public Property Get DefaultExt() As String
  DefaultExt = mDefaultExt
End Property
Public Property Let DefaultExt(StrExt As String)
  mDefaultExt = StrExt
End Property

'Public Property DereferenceLinks() As Boolean
Public Property Get DereferenceLinks() As Boolean       'OFN_NODEREFERENCELINKS
  DereferenceLinks = mDereferenceLinks
End Property
Public Property Let DereferenceLinks(BolVal As Boolean) 'OFN_NODEREFERENCELINKS
  mFlags = mFlags Or OFN_NODEREFERENCELINKS
  If BolVal Then
    mFlags = mFlags Xor OFN_NODEREFERENCELINKS
  End If
End Property
Private Sub mDereferenceLinks_OnLet(ByVal sender As Object)
Dim BolVal As Boolean
  'If sender Is Boolean Then
  BolVal = sender
  mFlags = mFlags Or OFN_NODEREFERENCELINKS
  If BolVal Then 'Achtung hier nicht Not !! weil: Not No 2*Verneinung
    mFlags = mFlags Xor OFN_NODEREFERENCELINKS
  End If
End Sub

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
'Private Sub mFilterIndex_OnLet(ByVal sender As Object)
'  '
'End Sub

'Ruft das Ausgangsverzeichnis ab, oder legt dieses fest
'Public Property InitialDirectory() As String
Public Property Get InitialDirectory() As String       'lpstrInitialDir
  'InitialDirectory = mInitialDirectory
  InitialDirectory = mInitialDirectory
End Property
Public Property Let InitialDirectory(StrVal As String) 'lpstrInitialDir
  mInitialDirectory = StrVal
End Property
'Private Sub mInitialDirectory_OnLet(ByVal sender As Object)
'  '
'End Sub

'MultiSelect hat nur der Opendialog nicht im Savedialog möglich
'ReadOnlyChecked hat nur der Opendialog nicht im Savedialog möglich

'ruft einen Wert ab, oder legt diesen fest, der angibt,
'ob das Dialogfeld das vor dem Schließen aktuelle Verzeichnis wiederherstellt,
'Public Property RestoreDirectory() As Boolean
Public Property Get RestoreDirectory() As Boolean       'OFN_NOCHANGEDIR
  RestoreDirectory = mRestoreDirectory
End Property
Public Property Let RestoreDirectory(BolVal As Boolean) 'OFN_NOCHANGEDIR
  mFlags = mFlags Or OFN_NOCHANGEDIR
  If Not BolVal Then
    mFlags = mFlags Xor OFN_NOCHANGEDIR
  End If
End Property
'Private Sub mRestoreDirectory_OnLet(ByVal sender As Object)
'  mFlags = mFlags Or OFN_NOCHANGEDIR
'  If Not sender Then
'    mFlags = mFlags Xor OFN_NOCHANGEDIR
'  End If
'End Sub

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

'ShowReadOnly hat nur der Opendialog nicht im Savedialog möglich
'Public Property Get ShowReadOnly() As Boolean       'OFN_HIDEREADONLY
'  Set ShowReadOnly = mShowReadOnly
'  'ShowReadOnly = mFlags And OFN_HIDEREADONLY
'End Property
''Public Property Set ShowReadOnly(BolVal As Boolean) 'OFN_HIDEREADONLY
''  Set mShowReadOnly = BolVal
''End Property
'Private Sub mShowReadOnly_OnLet(ByVal sender As Object)
'  mFlags = mFlags Or OFN_HIDEREADONLY
'  If sender Then
'    mFlags = mFlags Xor OFN_HIDEREADONLY
'  End If
'End Sub
  
'Public Property Title() As String
Public Property Get Title() As String       'lpstrTitle
  Title = mTitle
End Property
Public Property Let Title(StrVal As String) 'lpstrTitle
  mTitle = StrVal
End Property

'ob das Dialogfeld nur gültige Win32-Dateinamen akzeptiert
'Public Property ValidateNames() As Boolean
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
Dim OFN As OPENFILENAME
Dim hhwndOwner As Long
Dim mMaxLenFN As Long
Dim Buffer As String
'Dim Result As Long
Dim FN As String

TryE: On Error GoTo Catch
  
  If IsMissing(Frm) Then
    hhwndOwner = GetActiveWindow
  Else
    hhwndOwner = Frm.hwnd
  End If
  
  mMaxLenFN = 32767 - Len(mFileName)
  Buffer = mFileName & String$(mMaxLenFN, vbNullChar)
  
  With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = hhwndOwner
    .hInstance = App.hInstance
    .lpstrFilter = mFilter
    'lpstrCustomFilter
    'nMaxCustFilter
    .nFilterIndex = CLng(FilterIndex)
    .lpstrFile = Buffer
    .nMaxFile = Len(Buffer) '32767
    'lpstrFileTitle
    'nMaxFileTitle
    .lpstrInitialDir = mInitialDirectory
    .lpstrTitle = Title
    .flags = mFlags
    'nFileOffset
    'nFileExtension = 1
    If mAddExtension Then
      If Len(mDefaultExt) = 0 Then
        '
      End If
      .lpstrDefExt = mDefaultExt
    End If
    'lCustData
    'lpfnHook
    'lpTemplateName
  End With
  ShowDialog = GetSaveFileName(OFN)
  'Result = GetOpenFileNamePreview(OFN)
  'If Result <> 0 Then
    'man muß am String erkennen, ob nur eine oder mehrere Dateien ausgewählt wurden
    Call GetFNsinCol(OFN.lpstrFile)
    'ShowDialog = DialogResult.DialogResult_OK
  'Else
    'If mFlags And OFN_ALLOWMULTISELECT Then
    '  Call GetFNsinCol(OFN.lpstrFile)
    'End If
    'ShowDialog = DialogResult.DialogResultCancel
    'oder zuviele Dateien
  'End If
  Exit Function
Catch:
  If Err.Number = 32755 Then
    ShowDialog = vbAbort
  Else
    MsgBox CStr(Err.Number) + " " + Err.Description
  End If
End Function

