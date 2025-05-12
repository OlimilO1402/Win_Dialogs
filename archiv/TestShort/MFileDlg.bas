Attribute VB_Name = "MFileDlg"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_Value]
    End Enum
#End If

Private Type OpenFilename         ' x86  ' x64
    lStructSize       As Long     '   4  '   4
    hwndOwner         As LongPtr  '   4  '   8
    hInstance         As LongPtr  '   4  '   8
    lpstrFilter       As LongPtr  '   4  '   8
    lpstrCustomFilter As LongPtr  '   4  '   8
    nMaxCustFilter    As Long     '   4  '   4
    nFilterIndex      As Long     '   4  '   4
    lpstrFile         As LongPtr  '   4  '   8
    nMaxFile          As Long     '   4  '   4
    lpstrFileTitle    As LongPtr  '   4  '   8
    nMaxFileTitle     As Long     '   4  '   4
    lpstrInitialDir   As LongPtr  '   4  '   8
    lpstrTitle        As LongPtr  '   4  '   8
    Flags             As Long     '   4  '   4
    nFileOffset       As Integer  '   2  '   2
    nFileExtension    As Integer  '   2  '   2
    lpstrDefExt       As LongPtr  '   4  '   8
    lCustData         As Long     '   4  '   4
    lpfnHook          As LongPtr  '   4  '   8
    lpTemplateName    As LongPtr  '   4  '   8
End Type                   ' Sum: '  76  ' 120

Private Declare Function GetSaveFileNameW Lib "comdlg32" (SaveFilename As Any) As Long

' ################# Öffnen/Speichern-Dialog #######################
#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32" (OpenFilename As Any) As Long
    Private Declare PtrSafe Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As LongPtr, ByVal lpstrFile As LongPtr, ByVal nMaxFile As Long, ByVal lpstrInitialDir As LongPtr, ByVal lpstrDefExt As LongPtr, ByVal lpstrFilter As LongPtr, ByVal lpstrTitle As LongPtr) As Long 'Ptr
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function GetOpenFileNameW Lib "comdlg32" (OpenFilename As Any) As Long
    Private Declare Function GetFileNameFromBrowseW Lib "Shell32" Alias "#63" (ByVal hwndOwner As LongPtr, ByVal lpstrFile As LongPtr, ByVal nMaxFile As Long, ByVal lpstrInitialDir As LongPtr, ByVal lpstrDefExt As LongPtr, ByVal lpstrFilter As LongPtr, ByVal lpstrTitle As LongPtr) As Long 'Ptr
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#End If

Public Function OpenFile_ShowDialog(ByVal ahWndOwner As LongPtr, ByVal InitialDir As String, ByVal DefaultExt As String, ByVal Filter As String, ByVal Title As String, PathFileName_out As String) As VbMsgBoxResult
    Dim Buffer As String: Buffer = String(2048, vbNullChar)
    InitialDir = InitialDir & vbNullChar
    Filter = Replace(Filter, "|", vbNullChar) & vbNullChar
    Title = Title & vbNullChar
    Dim hr As Long: hr = GetFileNameFromBrowseW(ahWndOwner, StrPtr(Buffer), Len(Buffer), StrPtr(InitialDir), StrPtr(DefaultExt), StrPtr(Filter), StrPtr(Title))
    OpenFile_ShowDialog = IIf(hr, VbMsgBoxResult.vbOK, VbMsgBoxResult.vbCancel)
    If OpenFile_ShowDialog <> VbMsgBoxResult.vbOK Then Exit Function
    Dim l As Long: l = lstrlenW(StrPtr(Buffer))
    PathFileName_out = Left$(Buffer, l)
End Function

Public Function OpenFile_ShowDialog2(ByVal ahWndOwner As LongPtr, ByVal InitialDir As String, ByVal DefaultExt As String, ByVal Filter As String, ByVal Title As String, PathFileName_out As String) As VbMsgBoxResult
    ReDim OFN(0 To 18) As Long
    Dim i As Long: OFN(i) = 76                             'OFN.lStructSize
    i = i + 1:     OFN(i) = ahWndOwner                     'OFN.hwndOwner
    i = i + 1:     'OFN(i) = App.hInstance 'not necessary  'OFN.hInstance
    Filter = Replace(Filter, "|", vbNullChar) & vbNullChar '
    i = i + 1:     OFN(i) = StrPtr(Filter)                 'OFN.lpstrFilter
    i = i + 1:     'OFN(i) =                               'OFN.lpstrCustomFilter
    i = i + 1:     'OFN(i) =                               'OFN.nMaxCustFilter
    i = i + 1:     'OFN(i) =                               'OFN.nFilterIndex
    Dim Buffer As String: Buffer = String(2048, vbNullChar)
    i = i + 1:     OFN(i) = StrPtr(Buffer)                  'OFN.lpstrFile
    i = i + 1:     OFN(i) = 1024                           'OFN.nMaxFile
    i = i + 1:     'OFN(i) =                               'OFN.lpstrFileTitle
    i = i + 1:     'OFN(i) =                               'OFN.nMaxFileTitle
    Title = Title & vbNullChar
    i = i + 1:     OFN(i) = StrPtr(Title)
    InitialDir = InitialDir & vbNullChar
    i = i + 1:     OFN(i) = StrPtr(InitialDir)             'OFN.lpstrInitialDir
    
    Dim hr As Long
    hr = GetOpenFileNameW(OFN(0))
    
    OpenFile_ShowDialog2 = IIf(hr, VbMsgBoxResult.vbOK, VbMsgBoxResult.vbCancel)
    If OpenFile_ShowDialog2 <> VbMsgBoxResult.vbOK Then Exit Function
    Dim l As Long: l = lstrlenW(StrPtr(Buffer))
    PathFileName_out = Left$(Buffer, l)
End Function

Public Function OpenFile_ShowDialog3(ByVal ahWndOwner As LongPtr, ByVal InitialDir As String, ByVal DefaultExt As String, ByVal Filter As String, ByVal Title As String, PathFileName_out As String) As VbMsgBoxResult
    Dim Buffer As String: Buffer = String(2048, vbNullChar)
    Filter = Replace(Filter, "|", vbNullChar) & vbNullChar  '
    Title = Title & vbNullChar
    InitialDir = InitialDir & vbNullChar
    ReDim OFN(0 To 18) As Long
    OFN(0) = 76:                OFN(1) = ahWndOwner
    OFN(3) = StrPtr(Filter):    OFN(7) = StrPtr(Buffer):    OFN(8) = Len(Buffer)
    OFN(11) = StrPtr(Title):    OFN(12) = StrPtr(InitialDir)
    Dim hr As Long: hr = GetOpenFileNameW(OFN(0))
    OpenFile_ShowDialog3 = IIf(hr, VbMsgBoxResult.vbOK, VbMsgBoxResult.vbCancel)
    If OpenFile_ShowDialog3 <> VbMsgBoxResult.vbOK Then Exit Function
    Dim l As Long: l = lstrlenW(StrPtr(Buffer))
    PathFileName_out = Left$(Buffer, l)
End Function


