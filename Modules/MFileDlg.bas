Attribute VB_Name = "MFileDlg"
Option Explicit

'Private Type OPENFILENAMEW                 ' x86    ' x64
'    lStructSize       As Long              '   4  ' 0   '   4 + 4Padb ' 0
'    hwndOwner         As LongPtr ' Long    '   4  ' 1   '   8         ' 1
'    hInstance         As LongPtr ' Long    '   4  ' 2   '   8         ' 2
'    lpstrFilter       As LongPtr ' String  '   4  ' 3   '   8         ' 3
'    lpstrCustomFilter As LongPtr ' String  '   4  ' 4   '   8         ' 4
'    nMaxCustFilter    As Long              '   4  ' 5   '   4         ' 5
'    nFilterIndex      As Long              '   4  ' 6   '   4         '
'    lpstrFile         As LongPtr ' String  '   4  ' 7   '   8         ' 6
'    nMaxFile          As Long              '   4  ' 8   '   4 + 4Padb ' 7
'    lpstrFileTitle    As LongPtr ' String  '   4  ' 9   '   8         ' 8
'    nMaxFileTitle     As Long              '   4  '10   '   4 + 4Padb ' 9
'    lpstrInitialDir   As LongPtr ' String  '   4  '11   '   8         '10
'    lpstrTitle        As LongPtr ' String  '   4  '12   '   8         '11
'    flags             As Long              '   4  '13   '   4         '12
'    nFileOffset       As Integer           '   2  '14   '   2         '
'    nFileExtension    As Integer           '   2  '     '   2         '
'    lpstrDefExt       As LongPtr ' String  '   4  '15   '   8         '13
'    lCustData         As LongPtr ' Long    '   4  '16   '   8         '14
'    lpfnHook          As LongPtr ' Long    '   4  '17   '   8         '15
'    lpTemplateName    As LongPtr ' String  '   4  '18   '   8         '16
'End Type                               ' Sum: 76        ' 136

#If Win64 Then
    Private Const m_uOSFD   As Long = 16
    Private Const m_SizeByt As Long = (m_uOSFD + 1) * 8 '136
#Else
    Private Const m_uOSFD   As Long = 18
    Private Const m_SizeByt As Long = (m_uOSFD + 1) * 4 '76
#End If
Private m_OSFD(0 To m_uOSFD) As LongPtr 'in VBA it will be translated to LongLong automatically

' ################# Öffnen/Speichern-Dialog #######################
#If VBA7 Then
    'Private Declare PtrSafe Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As LongPtr, ByVal lpstrFile As LongPtr, ByVal nMaxFile As Long, ByVal lpstrInitialDir As LongPtr, ByVal lpstrDefExt As LongPtr, ByVal lpstrFilter As LongPtr, ByVal lpstrTitle As LongPtr) As Long 'Ptr
    Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32" (OpenFilename As Any) As Long
    Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32" (OpenFilename As Any) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#Else
    'Private Declare Function GetFileNameFromBrowseW Lib "Shell32" Alias "#63" (ByVal hwndOwner As LongPtr, ByVal lpstrFile As LongPtr, ByVal nMaxFile As Long, ByVal lpstrInitialDir As LongPtr, ByVal lpstrDefExt As LongPtr, ByVal lpstrFilter As LongPtr, ByVal lpstrTitle As LongPtr) As Long 'Ptr
    Private Declare Function GetOpenFileNameW Lib "comdlg32" (OpenFilename As Any) As Long
    Private Declare Function GetSaveFileNameW Lib "comdlg32" (SaveFilename As Any) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#End If

'just a little macro for VBA
'Sub Start()
'    UserForm1.Show
'End Sub
Public Function OpenFile_ShowDialog(ByVal ahWndOwner As LongPtr, ByVal InitialDir As String, ByVal DefaultExt As String, ByVal Filter As String, ByVal Title As String, PathFileName_out As String) As VbMsgBoxResult
    Dim Buffer As String
    Prepare ahWndOwner, InitialDir, DefaultExt, Filter, Title, Buffer
    Dim hr As Long: hr = GetOpenFileNameW(m_OSFD(0))
    OpenFile_ShowDialog = IIf(hr, VbMsgBoxResult.vbOK, VbMsgBoxResult.vbCancel)
    If OpenFile_ShowDialog <> VbMsgBoxResult.vbOK Then Exit Function
    Dim l As Long: l = lstrlenW(StrPtr(Buffer))
    PathFileName_out = Left$(Buffer, l)
End Function
Private Function GetAppHInstancePtr() As LongPtr
#If VBA7 And Win64 Then
    GetAppHInstancePtr = Excel.Application.hInstancePtr
#Else
    GetAppHInstancePtr = App.hInstance
#End If
End Function
Private Sub Prepare(ahWndOwner As LongPtr, InitialDir As String, DefaultExt As String, Filter As String, Title As String, Buffer As String)
    Buffer = String(2048, vbNullChar)
    Filter = Replace(Filter, "|", vbNullChar) & vbNullChar  '
    Title = Title & vbNullChar
    InitialDir = InitialDir & vbNullChar
    Dim i As Long
    m_OSFD(i) = m_SizeByt:      i = i + 1:    m_OSFD(i) = ahWndOwner:    i = i + 1:    m_OSFD(i) = GetAppHInstancePtr:    i = i + 1:    m_OSFD(i) = StrPtr(Filter):    i = i + 4
#If Win64 Then
    i = i - 1
#End If
    m_OSFD(i) = StrPtr(Buffer): i = i + 1:    m_OSFD(i) = Len(Buffer):   i = i + 3:    m_OSFD(i) = StrPtr(InitialDir):    i = i + 1:    m_OSFD(i) = StrPtr(Title):     i = i + 1
    m_OSFD(i) = 530436 'combination of flags
End Sub
Public Function SaveFile_ShowDialog(ByVal ahWndOwner As LongPtr, ByVal InitialDir As String, ByVal DefaultExt As String, ByVal Filter As String, ByVal Title As String, PathFileName_out As String) As VbMsgBoxResult
    Dim Buffer As String
    Prepare ahWndOwner, InitialDir, DefaultExt, Filter, Title, Buffer
    Dim hr As Long: hr = GetSaveFileNameW(m_OSFD(0))
    SaveFile_ShowDialog = IIf(hr, VbMsgBoxResult.vbOK, VbMsgBoxResult.vbCancel)
    If SaveFile_ShowDialog <> VbMsgBoxResult.vbOK Then Exit Function
    Dim l As Long: l = lstrlenW(StrPtr(Buffer))
    PathFileName_out = Left$(Buffer, l)
End Function

