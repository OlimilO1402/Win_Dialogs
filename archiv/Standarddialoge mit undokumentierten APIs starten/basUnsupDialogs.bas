Attribute VB_Name = "basUnsupDialogs"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

' Code by I.Runge (mastermind@ircastle.de)

' ################# Formatieren-Dialog #######################
Enum SHFD_CAPACITY
    SHFD_CAPACITY_DEFAULT = 0 ' standard Laufwerks-Kapazität
    SHFD_CAPACITY_360 = 3 ' 360KB, also nur für 5.25-Zoll-Laufwerke
    SHFD_CAPACITY_720 = 5 ' 7720KB, also nur für 3.5-Zoll-Laufwerke
End Enum

Enum SHFD_FORMAT
    SHFD_FORMAT_QUICK = 0 ' Schnell-Formatierung
    SHFD_FORMAT_FULL = 1 ' volle Formatierung
    SHFD_FORMAT_SYSONLY = 2 ' DOS-Startdiskette erstellen (nur Win95/98/ME)
End Enum

Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long


' ################# Ausführen-Dialog #######################
Const SHRD_NOMRU = &H2
Private Declare Function SHRunDialogA Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
    
Private Declare Function SHRunDialogW Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As Long, ByVal szPrompt As Long, ByVal uFlags As Long) As Long


' ################# Icon-Auswahl-Dialog #######################
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
    
Private Declare Function SHChangeIconDialogA Lib "shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long
    
Private Declare Function SHChangeIconDialogW Lib "shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As Long, ByVal Reserved As Long, lpIconIndex As Long) As Long

' ################# Windows-Beenden-Dialog #######################
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal lSelOption As Long) As Long


' ################# Öffnen/Speichern-Dialog #######################
Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
    
Private Declare Function GetFileNameFromBrowseA Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, ByVal lpstrInitialDir As String, ByVal lpstrDefExt As String, ByVal lpstrFilter As String, ByVal lpstrTitle As String) As Long

Public Function ShowOpenDlg(ByVal Owner As Form, Optional ByVal InitialDir As String, Optional ByVal Filter As String, Optional ByVal DefaultExtension As String, Optional ByVal DlgTitle As String) As String
    
    Dim sBuf As String
    
    InitialDir = IIf(IsMissing(InitialDir), "", InitialDir)
    Filter = IIf(IsMissing(Filter), "Alle Dateien|*.*", Replace(Filter, "|", vbNullChar)) & vbNullChar
    DefaultExtension = IIf(IsMissing(DefaultExtension), "", DefaultExtension)
    DlgTitle = IIf(IsMissing(DlgTitle), "Datei wählen", DlgTitle)
    
    sBuf = Space(256)
    If IsWinNT Then
        Call GetFileNameFromBrowseW(Owner.hWnd, StrPtr(sBuf), Len(sBuf), StrPtr(InitialDir), StrPtr(DefaultExtension), StrPtr(Filter), StrPtr(DlgTitle))
    Else
        Call GetFileNameFromBrowseA(Owner.hWnd, sBuf, Len(sBuf), InitialDir, DefaultExtension, Filter, DlgTitle)
    End If
    
    ShowOpenDlg = Trim(sBuf)
        
End Function

Public Sub ShowShutDownDlg()
    Dim rv As Long: rv = SHShutDownDialog(0&)
End Sub

Private Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function ShowPicIconDlg(ByVal Owner As Form, ByRef rsIconFile As String, ByRef rlIconIndex As Long) As Boolean
    Dim fn As String * 260, l As Long, lngResult As Long
    
    fn = rsIconFile & vbNullChar
    '"C:\Windows\System32\Shell32.dll"
    If IsWinNT Then
        'Unicode
        lngResult = SHChangeIconDialogW(Owner.hWnd, StrPtr(fn), l, rlIconIndex)
    Else
        'ANSI
        lngResult = SHChangeIconDialogA(Owner.hWnd, fn, l, rlIconIndex)
    End If
    
    ShowPicIconDlg = (lngResult <> 0)
    If ShowPicIconDlg Then rsIconFile = Left(fn, InStr(1, fn, vbNullChar) - 1)
    
End Function

Public Sub ShowRunDlg(ByVal Owner As Form, _
    Optional ByVal DontShowLastFileName As Boolean = False)
    
    Const DlgTitle = "Ausführen"
    Const DlgText = "Geben Sie den Namen des Programms, Ordners oder Dokuments an, " & "das bzw. der geöffnet werden soll."
    
    If Not IsWinNT Then
        SHRunDialogA Owner.hWnd, 0&, 0&, DlgTitle, DlgText, IIf(DontShowLastFileName, SHRD_NOMRU, 0&)
    Else
        SHRunDialogW Owner.hWnd, 0&, 0&, StrPtr(DlgTitle), StrPtr(DlgText), IIf(DontShowLastFileName, SHRD_NOMRU, 0&)
    End If
End Sub

Public Sub ShowFormatDriveDlg(ByVal Owner As Form, Optional ByVal DriveLetter As String = "A", Optional ByVal Capacity As SHFD_CAPACITY, Optional ByVal FormatMode As SHFD_FORMAT)
    
    'iDrive = Nummer des Laufwerks (A=0, B=1, C=2, usw.)
    SHFormatDrive Owner.hWnd, Asc(UCase(DriveLetter)) - 65, Capacity, FormatMode
End Sub

