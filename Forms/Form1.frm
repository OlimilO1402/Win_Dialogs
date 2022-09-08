VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   Caption         =   "WinDialogs"
   ClientHeight    =   6015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "Test MyFontDialog"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestMessageBox 
      Caption         =   "Test MessageBox"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton BtnAllDialogs 
      Caption         =   "All Dialogs (old)"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "locale Folders"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Printer"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Computer"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Folders"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton BtnFolderBrowser 
      Caption         =   "FolderBrowser"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   5055
   End
   Begin VB.ComboBox CmbSpecialFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   4320
      Width           =   5055
   End
   Begin VB.CheckBox ChkShowEditBox 
      Caption         =   "ShowEditBox"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.CheckBox ChkShowNewFolderButton 
      Caption         =   "ShowNewFolderButton"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   3840
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.TextBox TxtSelectedPath 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   5055
   End
   Begin VB.CheckBox ChkSelectedPath 
      Caption         =   "Use the path above as the starting folder if possible"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Value           =   1  'Aktiviert
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":1782
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Use FolderBrowserDialog as special dialog fo searching ..."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Label LblFBD 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   75
   End
   Begin VB.Label LblCD 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "    "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label LblFD 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "AaBbYyZz"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label LblSFD 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   75
   End
   Begin VB.Label LblOFD 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFilePrinter 
         Caption         =   "Printer"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditColorChoose 
         Caption         =   "Select &Color..."
      End
      Begin VB.Menu mnuEditFontChoose 
         Caption         =   "Select &Font..."
      End
      Begin VB.Menu mnuEditFolderChoose 
         Caption         =   "Select Folder (prefered)"
      End
      Begin VB.Menu mnuEditFolderSelectOFD 
         Caption         =   "Select Folder (simple OpenFiledialog)"
      End
      Begin VB.Menu mnuEditPathChoose 
         Caption         =   "Select Folder (old FolderBrowser)"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptionUseOldComDlg 
         Caption         =   "Use old CommonDialog-control"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CD  As New ColorDialog

Private Sub Command5_Click()
    Dim FD As New MyFontDialog
    'MsgBox Hex(FD.Options)
    Set FD.Font = LblFD.Font
    'FD.AllowVectorFonts = False
    'FD.AllowVerticalFonts = False
    'FD.FixedPitchOnly = True
    FD.ShowColor = True
    'FD.ShowHelp = True
    FD.Color = LblFD.ForeColor
    If FD.ShowDialog(Me.hwnd) = vbOK Then
        Set LblFD.Font = FD.Font
        LblFD.ForeColor = FD.Color
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & MApp.Version
    mnuFilePrinter.Visible = False
    PrepareSpecialFolder
End Sub

Private Sub BtnAllDialogs_Click()
    Me.CommonDialog.ShowColor
    Me.CommonDialog.ShowFont
    'Me.CommonDialog.ShowHelp
    Me.CommonDialog.ShowOpen
    Me.CommonDialog.ShowPrinter 'still missing here
    Me.CommonDialog.ShowSave
End Sub

Private Sub BtnTestMessageBox_Click()
    Dim hr As Long
    'hr = MessageBox(Me.hwnd, "Test", "Test", 1)
    'hr = MessageBox(Me.hwnd, ByVal "Test", ByVal "Test", 1)
    'hr = MessageBox(Me.hwnd, ByVal StrPtr("Test"), ByVal StrPtr("Test"), 1)
    
    Dim sTxt As String
    Dim sCap As String
    Dim i As Long
    Dim sStr As Long
    Dim sEnd As Long
    
    sTxt = "": sCap = ""
    sStr = 65: sEnd = sStr + 25
    For i = sStr To sEnd
        sTxt = sTxt & ChrW(i)
        sCap = sCap & ChrW(i + 32)
    Next
    
    'hr = MessageBox(Me.hwnd, ByVal StrPtr(sTxt), ByVal StrPtr(sCap), 1)
    hr = MsgBox(sTxt, , sCap)
    MsgBox MWin.LastMsgBoxResult
    
    sTxt = "": sCap = ""
    sStr = 913: sEnd = sStr + 25
    For i = sStr To sEnd
        sTxt = sTxt & ChrW(i)
        sCap = sCap & ChrW(i + 32)
    Next
    hr = MsgBox(sTxt, , sCap)
    MsgBox MWin.LastMsgBoxResult
    
    hr = MsgBox(sTxt, vbCritical Or vbMsgBoxHelpButton Or vbMsgBoxRight Or vbMsgBoxRtlReading Or vbYesNoCancel Or vbDefaultButton4, sCap)
    MsgBox MWin.LastMsgBoxResult
    
    With New MessageBox
        .Caption = sCap
        .Text = sTxt
        '.LanguageID =
        .HIcon = Me.Icon.Handle
        .MsgBoxFncType = vbIndirect
        .Style = vbCancelTryContinue Or vbInformation Or vbMsgBoxHelpButton Or vbDefaultButton4
        hr = .Show
        MsgBox .Result_ToStr(hr)
    End With
    MsgBox MWin.HelpInfo_ToStr
End Sub

Private Sub mnuEditFolderSelectOFD_Click()
    With New OpenFileDialog
        'Dim sfs As String: sfs = "Folder Selection"
        .Title = "Select a Folder"
        .InitialDirectory = LblFBD.Caption
        .FileName = "Folder Selection"
        .ValidateNames = False
        .CheckFileExists = False
        .CheckPathExists = True
        If .ShowDialog(Me) = vbOK Then
            Dim FNm As String: FNm = .FileName
            Dim pos As Long: pos = InStrRev(FNm, "\")
            If pos > 3 Then
                FNm = Left(FNm, pos)
                LblFBD.Caption = FNm
            End If
        End If
    End With
End Sub

Private Sub mnuFileOpen_Click()
    Dim FNm As String
    If mnuOptionUseOldComDlg.Checked Then FNm = FileOpenOld Else FNm = FileOpenNew
    If Len(FNm) Then
        MsgBox FNm
        LblOFD.Caption = FNm
    End If
End Sub
Private Function FileOpenNew() As String
    With New OpenFileDialog
        .Filter = MApp.FileExtFilter
        .CheckFileExists = False
        .CheckPathExists = False
        .DefaultExt = ".htm"
        .ShowReadOnly = True
        .AddExtension = False
        .MultiSelect = True
        If .ShowDialog = vbOK Then
            FileOpenNew = .FileName
        End If
        Dim FNm
        Dim s As String
        For Each FNm In .FileNames
            s = s & FNm & vbCrLf
        Next
        MsgBox s
    End With
End Function
Private Function FileOpenOld() As String
Try: On Error GoTo Catch
    With Me.CommonDialog
        .Filter = MApp.FileExtFilter
        .Flags = .Flags Or FileOpenConstants.cdlOFNFileMustExist
        .Flags = .Flags Or FileOpenConstants.cdlOFNPathMustExist
        .DefaultExt = ".htm"
        .CancelError = True
        .Flags = .Flags Or FileOpenConstants.cdlOFNAllowMultiselect
        .Flags = .Flags Or FileOpenConstants.cdlOFNReadOnly
        .ShowOpen
        FileOpenOld = .FileName
    End With
Catch:
    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
        MComDlgCtrl.MessCommonDlgError Err.Number
    End If
End Function

Private Sub mnuFileSaveAs_Click()
    Dim FNm As String
    If mnuOptionUseOldComDlg.Checked Then FNm = FileSaveOld Else FNm = FileSaveNew
    If Len(FNm) Then LblSFD.Caption = FNm
End Sub
Private Function FileSaveNew() As String
    With New SaveFileDialog
        .Filter = MApp.FileExtFilter
        'Debug.Print .Filter
        '.ShowHelp = True
        If .ShowDialog = vbCancel Then Exit Function
        FileSaveNew = .FileName
    End With
End Function
Private Function FileSaveOld() As String
Try: On Error GoTo Catch
    With Me.CommonDialog
        .Filter = MApp.FileExtFilter
        .Flags = .Flags Or FileOpenConstants.cdlOFNFileMustExist
        .Flags = .Flags Or FileOpenConstants.cdlOFNPathMustExist
        .DefaultExt = ".htm"
        .CancelError = True
        .Flags = .Flags Or FileOpenConstants.cdlOFNReadOnly
        .ShowOpen
        FileSaveOld = .FileName
    End With
Catch:
    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
        MComDlgCtrl.MessCommonDlgError Err.Number
    End If
End Function

Private Sub mnuFilePrinter_Click()
    Dim PNm As String
    If mnuOptionUseOldComDlg.Checked Then PNm = FilePrinterOld Else PNm = FilePrinterNew
    If Len(PNm) Then MsgBox PNm
End Sub

Private Function FilePrinterNew() As String
    FilePrinterNew = Printer.DeviceName
End Function
Private Function FilePrinterOld() As String
Try: On Error GoTo Catch
    With Me.CommonDialog
        .CancelError = True
        .ShowPrinter
        FilePrinterOld = Printer.DeviceName
    End With
Catch:
    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
        MComDlgCtrl.MessCommonDlgError Err.Number
    End If
End Function

'--------------------------------------------------
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'==================================================
Private Sub mnuEditColorChoose_Click()
    Dim col As Long
    If mnuOptionUseOldComDlg.Checked Then col = ColorChooseOld Else col = ColorChooseNew
    If col = -1 Then Exit Sub
    LblCD.BackColor = col
End Sub
Private Function ColorChooseNew() As Long
    ColorChooseNew = -1
    With CD
        .Color = LblCD.BackColor
        .SolidColorOnly = True
        If .ShowDialog = vbCancel Then Exit Function
        ColorChooseNew = .Color
    End With
End Function
Private Function ColorChooseOld() As Long
    ColorChooseOld = -1
Try: On Error GoTo Catch
    With CommonDialog
        .Color = LblCD.BackColor
        .CancelError = True
        .ShowColor
        ColorChooseOld = .Color
    End With
Catch:
    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
        MComDlgCtrl.MessCommonDlgError Err.Number
    End If
End Function

Private Sub mnuEditFontChoose_Click()
    Dim f As StdFont: Set f = LblFD.Font
    Dim C As Long:        C = LblFD.ForeColor
    If mnuOptionUseOldComDlg.Checked Then Set f = FontDialogOld(f, C) Else Set f = FontDialogNew(f, C)
    Set LblFD.Font = f
    LblFD.ForeColor = C
End Sub

Private Function FontDialogNew(Font_in As StdFont, ByRef Color_inout As Long) As StdFont
    With New FontDialog
        Set .Font = Font_in
        .Color = Color_inout
        If .ShowDialog = vbCancel Then Exit Function
        Set FontDialogNew = .Font
        Color_inout = .Color
    End With
End Function
Private Function FontDialogOld(Font_in As StdFont, ByRef Color_inout As Long) As StdFont
Try: On Error GoTo Catch
    Dim Font As StdFont
    With CommonDialog
        .CancelError = True
        .Color = Color_inout
        .FontName = Font_in.Name
        .FontSize = Font_in.Size
        .FontBold = Font_in.Bold
        .FontItalic = Font_in.Italic
        .FontUnderline = Font_in.Underline
        .FontStrikethru = Font_in.Strikethrough
        .ShowFont
    End With
Catch:
    FontDialogOld = Font_in
    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
        MComDlgCtrl.MessCommonDlgError Err.Number
    End If
End Function

Private Sub mnuEditFolderChoose_Click()
    With New OpenFolderDialog
        '.Title = "Select a folder"
        
        'it is not needed to set the last folder
        'the dialog already knows it
        '.Folder = LblFBD.Caption
        
        ' If you do not set the property Folder, the default Folder is "Dieser PC"
        ' you may also try "C:", "C:\" (=drive C:), "C:\Downloads", "C:\Downloads\" (= Ordner Downloads auf Laufwerk C:)
        ' "::{645FF040-5081-101B-9F08-00AA002F954E}" = Paperbin  Papierkorb
        ' "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" = Network   Netzwerk
        ' "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}" = Libraries Bibliotheken
        '.Folder =
        
        If .ShowDialog(Me.hwnd) = vbOK Then
            LblFBD.Caption = .Folder
        Else
            
        End If
    End With
End Sub

Private Sub mnuEditPathChoose_Click()
    With New FolderBrowserDialog
        .Description = "Please select a folder:"
        .ShowNewFolderButton = True
        If .ShowDialog(Me) = vbOK Then
            LblFBD.Caption = .SelectedPath
        End If
    End With
End Sub

'--------------------------------------
Private Sub mnuOptionUseOldComDlg_Click()
    mnuOptionUseOldComDlg.Checked = Not mnuOptionUseOldComDlg.Checked
    Dim bUseOldComDlg As Boolean: bUseOldComDlg = mnuOptionUseOldComDlg.Checked
    mnuFilePrinter.Visible = bUseOldComDlg
End Sub
'--------------------------------------
Private Sub mnuHelpInfo_Click()
    Dim s As String
    With App
        s = s & .CompanyName & " " & .ProductName & vbCrLf
        s = s & .FileDescription & vbCrLf
        s = s & "Version: " & MApp.Version
    End With
    MsgBox s
End Sub

' v ############################## v ' based on SHBrowseForFolder deprecated ' v ############################## v '
Private Sub PrepareSpecialFolder()
    TxtSelectedPath.Text = App.Path
    With CmbSpecialFolder
        Call .AddItem("SpecialFolder_Desktop"):     .ItemData(.NewIndex) = SpecialFolder_Desktop
        Call .AddItem("CSIDL_INTERNET"):            .ItemData(.NewIndex) = CSIDL_INTERNET
        Call .AddItem("SpecialFolder_Programs"):    .ItemData(.NewIndex) = SpecialFolder_Programs
        Call .AddItem("CSIDL_CONTROLS"):            .ItemData(.NewIndex) = CSIDL_CONTROLS
        Call .AddItem("CSIDL_PRINTERS"):            .ItemData(.NewIndex) = CSIDL_PRINTERS
        Call .AddItem("SpecialFolder_Personal"):    .ItemData(.NewIndex) = SpecialFolder_Personal
        Call .AddItem("SpecialFolder_Favorites"):   .ItemData(.NewIndex) = SpecialFolder_Favorites
        Call .AddItem("SpecialFolder_Startup"):     .ItemData(.NewIndex) = SpecialFolder_Startup
        Call .AddItem("SpecialFolder_Recent"):      .ItemData(.NewIndex) = SpecialFolder_Recent
        Call .AddItem("SpecialFolder_SendTo"):      .ItemData(.NewIndex) = SpecialFolder_SendTo
        Call .AddItem("CSIDL_BITBUCKET"):           .ItemData(.NewIndex) = CSIDL_BITBUCKET
        Call .AddItem("SpecialFolder_StartMenu"):   .ItemData(.NewIndex) = SpecialFolder_StartMenu
    '&HC ??
        Call .AddItem("SpecialFolder_MyMusic"):     .ItemData(.NewIndex) = SpecialFolder_MyMusic
    '&HE, &HF ??
        Call .AddItem("SpecialFolder_DesktopDirectory")
                                         .ItemData(.NewIndex) = SpecialFolder_DesktopDirectory
        Call .AddItem("SpecialFolder_MyComputer"):  .ItemData(.NewIndex) = SpecialFolder_MyComputer
        Call .AddItem("CSIDL_NETWORK"):             .ItemData(.NewIndex) = CSIDL_NETWORK
        'Hood = Umgebung
        Call .AddItem("CSIDL_NETHOOD"):             .ItemData(.NewIndex) = CSIDL_NETHOOD
        Call .AddItem("CSIDL_FONTS"):               .ItemData(.NewIndex) = CSIDL_FONTS
        Call .AddItem("SpecialFolder_Templates"):   .ItemData(.NewIndex) = SpecialFolder_Templates
        Call .AddItem("CSIDL_COMMON_STARTMENU"):    .ItemData(.NewIndex) = CSIDL_COMMON_STARTMENU
        Call .AddItem("CSIDL_COMMON_PROGRAMS"):     .ItemData(.NewIndex) = CSIDL_COMMON_PROGRAMS
        Call .AddItem("CSIDL_COMMON_STARTUP"):      .ItemData(.NewIndex) = CSIDL_COMMON_STARTUP
        Call .AddItem("CSIDL_COMMON_DESKTOPDIRECTORY")
                                         .ItemData(.NewIndex) = CSIDL_COMMON_DESKTOPDIRECTORY
        Call .AddItem("SpecialFolder_ApplicationData")
                                         .ItemData(.NewIndex) = SpecialFolder_ApplicationData
        Call .AddItem("CSIDL_PRINTHOOD"):           .ItemData(.NewIndex) = CSIDL_PRINTHOOD
        Call .AddItem("SpecialFolder_LocalApplicationData")
                                         .ItemData(.NewIndex) = SpecialFolder_LocalApplicationData
        Call .AddItem("CSIDL_ALTSTARTUP"):          .ItemData(.NewIndex) = CSIDL_ALTSTARTUP
        Call .AddItem("CSIDL_COMMON_ALTSTARTUP"):   .ItemData(.NewIndex) = CSIDL_COMMON_ALTSTARTUP
        Call .AddItem("CSIDL_COMMON_FAVORITES"):    .ItemData(.NewIndex) = CSIDL_COMMON_FAVORITES
        Call .AddItem("SpecialFolder_InternetCache")
                                         .ItemData(.NewIndex) = SpecialFolder_InternetCache
        Call .AddItem("SpecialFolder_Cookies"):     .ItemData(.NewIndex) = SpecialFolder_Cookies
        Call .AddItem("SpecialFolder_History"):     .ItemData(.NewIndex) = SpecialFolder_History
        Call .AddItem("SpecialFolder_CommonApplicationData")
                                         .ItemData(.NewIndex) = SpecialFolder_CommonApplicationData
        Call .AddItem("CSIDL_WINDOWS"):             .ItemData(.NewIndex) = CSIDL_WINDOWS
        Call .AddItem("SpecialFolder_System"):      .ItemData(.NewIndex) = SpecialFolder_System
        Call .AddItem("SpecialFolder_ProgramFiles")
                                         .ItemData(.NewIndex) = SpecialFolder_ProgramFiles
        Call .AddItem("SpecialFolder_MyPictures"):  .ItemData(.NewIndex) = SpecialFolder_MyPictures
        Call .AddItem("CSIDL_PROFILE"):             .ItemData(.NewIndex) = CSIDL_PROFILE
        Call .AddItem("CSIDL_SYSTEMX86"):           .ItemData(.NewIndex) = CSIDL_SYSTEMX86
        Call .AddItem("CSIDL_PROGRAM_FILESX86"):    .ItemData(.NewIndex) = CSIDL_PROGRAM_FILESX86
        Call .AddItem("SpecialFolder_CommonProgramFiles")
                                         .ItemData(.NewIndex) = SpecialFolder_CommonProgramFiles
        Call .AddItem("CSIDL_PROGRAM_FILES_COMMONX86")
                                         .ItemData(.NewIndex) = CSIDL_PROGRAM_FILES_COMMONX86
        Call .AddItem("CSIDL_COMMON_TEMPLATES"):    .ItemData(.NewIndex) = CSIDL_COMMON_TEMPLATES
        Call .AddItem("CSIDL_COMMON_DOCUMENTS"):    .ItemData(.NewIndex) = CSIDL_COMMON_DOCUMENTS
        Call .AddItem("CSIDL_COMMON_ADMINTOOLS"):   .ItemData(.NewIndex) = CSIDL_COMMON_ADMINTOOLS
        Call .AddItem("CSIDL_ADMINTOOLS"):          .ItemData(.NewIndex) = CSIDL_ADMINTOOLS
        Call .AddItem("CSIDL_CONNECTIONS"):         .ItemData(.NewIndex) = CSIDL_CONNECTIONS
        Call .AddItem("CSIDL_FLAG_DONT_VERIFY"):    .ItemData(.NewIndex) = CSIDL_FLAG_DONT_VERIFY
        Call .AddItem("CSIDL_FLAG_CREATE"):         .ItemData(.NewIndex) = CSIDL_FLAG_CREATE
        Call .AddItem("CSIDL_FLAG_MASK"):           .ItemData(.NewIndex) = CSIDL_FLAG_MASK
        Call .AddItem("CSIDL_FLAG_PFTI_TRACKTARGET")
                                         .ItemData(.NewIndex) = CSIDL_FLAG_PFTI_TRACKTARGET
        '.Text = "SpecialFolder_Desktop"
        .ListIndex = 0
    End With
    
End Sub

Private Sub BtnFolderBrowser_Click()
    Call ShowFBD(CmbSpecialFolder.ItemData(CmbSpecialFolder.ListIndex))
End Sub

Private Sub Command1_Click()
    Call ShowFBD(SpecialFolder_MyComputer)
End Sub

Private Sub Command2_Click()
    Call ShowFBD(CSIDL_NETWORK)
End Sub

Private Sub Command3_Click()
    Call ShowFBD(CSIDL_PRINTERS)
End Sub

Private Sub Command4_Click()
    Call ShowFBD(SpecialFolder_Personal)
End Sub

Private Sub ShowFBD(spf As Environment_SpecialFolder)
    With New FolderBrowserDialog 'FBD
        .RootFolder = spf
        Select Case spf
        Case SpecialFolder_MyComputer
                              .Flags = .Flags Or BIF_RETURNONLYFSDIRS
        Case CSIDL_NETWORK
                              .Flags = 0 'set to 0 before!
                              .Flags = .Flags Or BIF_BROWSEFORCOMPUTER
        Case CSIDL_PRINTERS
                              .Flags = .Flags Or BIF_BROWSEFORPRINTER
        Case SpecialFolder_Personal
                              '.Flags = .Flags Or BIF_DONTGOBELOWDOMAIN
                              .Flags = 0 'set to 0 before!
                              .Flags = .Flags Or BIF_RETURNFSANCESTORS
        End Select
        
        If ChkShowEditBox.Value = vbChecked Then
            .Flags = .Flags Or BIF_EDITBOX
        End If
        If Me.ChkShowNewFolderButton = vbUnchecked Then
            .Flags = .Flags Or BIF_DONTSHOWNEWFOLDERBUTTON
        End If
        'maximal 3 Zeilen Beschreibungstext
        .Description = "Hier sollte ein Hinweis stehen für den Benutzer was er hier tun soll. " & _
                       "In maximal 3 Zeilen erklärt. 12345 67890 12345 67890 12345 67890 " & _
                       "12345 67890 12345 67890 12345 67890 12345!"
        If (ChkSelectedPath.Value = vbChecked) And (Len(TxtSelectedPath.Text) > 0) Then
            .SelectedPath = TxtSelectedPath.Text
        End If
        If .ShowDialog(Me) = vbOK Then
            Dim s As String: s = .SelectedPath
            TxtSelectedPath.Text = s
            Dim t As String: t = TxtSelectedPath.Text
            Dim p As Long: p = InStr(1, t, "?")
            If p > 0 Then
                MsgBox s
            End If
        End If
    End With

End Sub
