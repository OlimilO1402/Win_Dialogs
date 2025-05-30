VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "WinDialogs"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox CommonDialog 
      Height          =   495
      Left            =   4800
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton BtnOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestTaskDialog 
      Caption         =   "Test TaskDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test MyFontDialog"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   2415
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
      Caption         =   $"FMain.frx":1782
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
      Width           =   735
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
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu mnuFilePrinter 
         Caption         =   "Print"
         Begin VB.Menu mnuFilePrinter1 
            Caption         =   "PrintDialog..."
         End
         Begin VB.Menu mnuFilePrinter2 
            Caption         =   "PrintDialogEx..."
         End
         Begin VB.Menu mnuFilePrinter3 
            Caption         =   "PrintDlgWinUI..."
         End
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
      Begin VB.Menu mnuEditFindReplace 
         Caption         =   "Find Replace..."
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptionUseUndocumShell32 
         Caption         =   "Use undocumented Shell32.dll functions"
      End
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
Private m_FR As FindReplaceDialog
Private m_StringToSearchIn As String
Private m_FindWhat As String


Private Sub BtnTestTaskDialog_Click()
    Dim vlg As String: vlg = ", we make it very long to see what happens if it is too long, maybe there are line breaks . . ."
    Dim tit As String: tit = "This is the title" & vlg
    Dim ins As String: ins = "This are the instructions" & vlg
    Dim con As String: con = "This is the content" & vlg
    Dim tdret As VbMsgBoxResult: tdret = MApp.TaskDialog(tit, ins, con, ETaskDialogIcon.tdIconWarning, tdButtonOK Or tdButtonCancel Or tdButtonClose).ShowDialog(Me)
    MsgBox "tdret: " & MWin.DialogResult_ToStr(tdret)
End Sub

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

Public Function SelectPrinter(ByVal PrinterName As String) As Printer
    Dim i As Long
    For i = 0 To Printers.Count - 1
        If UCase(Printers(i).DeviceName) = UCase(PrinterName) Then 'e.g.: "Microsoft Print to PDF"
            Set SelectPrinter = Printers(i)
            'Set Printer = SelectPrinter 'Printers(i)
            Exit For
        End If
    Next
End Function

Private Sub BtnOpen_Click()
    Dim PFN As String
    If MFileDlg.OpenFile_ShowDialog(Me.hwnd, App.Path, ".txt", "Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*", "File Open", PFN) = vbCancel Then Exit Sub
    Debug.Print PFN
End Sub

Private Sub BtnSave_Click()
    Dim PFN As String
    If MFileDlg.SaveFile_ShowDialog(Me.hwnd, App.Path, ".txt", "Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*", "File Save", PFN) = vbCancel Then Exit Sub
    Debug.Print PFN
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & MApp.Version
    'mnuFilePrinter.Visible = False
    PrepareSpecialFolder
End Sub

Private Sub BtnAllDialogs_Click()
'    Me.CommonDialog.ShowColor
'    Me.CommonDialog.ShowFont
'    'Me.CommonDialog.ShowHelp
'    Me.CommonDialog.ShowOpen
'    Me.CommonDialog.ShowPrinter 'still missing here
'    Me.CommonDialog.ShowSave
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

Private Sub mnuEditFindReplace_Click()
    If m_FR Is Nothing Then
        Set m_FR = New FindReplaceDialog
        m_FR.IsReplaceDlg = False
    End If
    m_StringToSearchIn = "Quick brown fos jumps over the lazy dog"
    m_FR.FindWhat = "over"
    m_FR.ReplaceWith = "   o v e r   "
    m_FR.MatchCase = True
    m_FR.ShowDialog Me
    MsgBox m_FR.LastError
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
    If mnuOptionUseUndocumShell32.Checked Then
        Dim Filter  As String: Filter = "Textfiles [*.txt]|*.txt|All Files [*.*]|*.*"
        Dim InitDir As String: InitDir = App.Path
        Dim DefExt  As String: DefExt = ".txt"
        Dim Title   As String: Title = "Open"
        Dim PathFileName As String
        'If MShell32U.OpenFile_ShowDialog(Me.hwnd, InitDir, DefExt, Filter, Title, PathFileName) = vbOK Then
        If MShell32U.SaveFile_ShowDialog(Me.hwnd, InitDir, DefExt, Filter, Title, PathFileName) = vbOK Then
            FNm = PathFileName
        End If
    Else
        If mnuOptionUseOldComDlg.Checked Then FNm = FileOpenOld Else FNm = FileOpenNew
    End If
    If Len(FNm) Then
        MsgBox FNm
        LblOFD.Caption = FNm
    End If
End Sub
Private Function FileOpenNew() As String
    With New OpenFileDialog
        '.Filter = MApp.FileExtFilter
        '.CheckFileExists = False
        '.CheckPathExists = False
        '.DefaultExt = ".htm"
        '.ShowReadOnly = True
        '.AddExtension = False
        '.MultiSelect = True
        If .ShowDialog = vbOK Then
            FileOpenNew = .FileName
        End If
        If .FileNames.Count > 1 Then
            Dim FNm
            Dim s As String
            For Each FNm In .FileNames
                s = s & FNm & vbCrLf
            Next
            MsgBox s
        End If
    End With
End Function
Private Function FileOpenOld() As String
'Try: On Error GoTo Catch
'    With Me.CommonDialog
'        .Filter = MApp.FileExtFilter
'        .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
'        .flags = .flags Or FileOpenConstants.cdlOFNPathMustExist
'        .DefaultExt = ".htm"
'        .CancelError = True
'        .flags = .flags Or FileOpenConstants.cdlOFNAllowMultiselect
'        .flags = .flags Or FileOpenConstants.cdlOFNReadOnly
'        .ShowOpen
'        FileOpenOld = .FileName
'    End With
'Catch:
'    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
'        MComDlgCtrl.MessCommonDlgError Err.Number
'    End If
End Function

Private Sub mnuFilePageSetup_Click()
    If mnuOptionUseOldComDlg.Checked Then
        ShowPageSetupdialogOld
    Else
        ShowPageSetupdialogNew
    End If
End Sub

Private Sub ShowPageSetupdialogOld()
'Try: On Error GoTo Catch
'    With Me.CommonDialog
'        .flags = .flags Or PrinterConstants.cdlPDPrintSetup
'        .ShowPrinter
'        MsgBox Printer.DeviceName
'    End With
'Catch:
'    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
'        MComDlgCtrl.MessCommonDlgError Err.Number
'    End If
End Sub

Private Sub ShowPageSetupdialogNew()
    Dim psd As New PageSetupDialog
    With psd
        If .ShowDialog = vbCancel Then Exit Sub
        MsgBox "Driver, Device, Output: " & vbCrLf & psd.DriverName & ", " & psd.DeviceName & ", " & psd.OutputName & vbCrLf & _
               "w*h: " & psd.PaperSizeWidth & " " & psd.PaperSizeHeight & vbCrLf & _
               "margin-l,r,t,b=" & psd.MarginsLeft & ", " & psd.MarginsRight & ", " & psd.MarginsTop & ", " & psd.MarginsBottom & vbCrLf & _
               "marginMin-l,r,t,b=" & psd.MinMarginsLeft & ", " & psd.MinMarginsRight & ", " & psd.MinMarginsTop & ", " & psd.MinMarginsBottom
               
               'Papiergr��e
               'Papierausrichtung
               'Quelle, Schacht
    End With
End Sub


Private Sub mnuFileSaveAs_Click()
    Dim FNm As String
    If mnuOptionUseUndocumShell32.Checked Then
        Dim Filter  As String: Filter = "Textfiles [*.txt]|*.txt|All Files [*.*]|*.*"
        Dim InitDir As String: InitDir = App.Path
        Dim DefExt  As String: DefExt = ".txt"
        Dim Title   As String: Title = "Save As..."
        Dim PathFileName As String
        'If MShell32U.OpenFile_ShowDialog(Me.hwnd, InitDir, DefExt, Filter, Title, PathFileName) = vbOK Then
        If MShell32U.SaveFile_ShowDialog(Me.hwnd, InitDir, DefExt, Filter, Title, PathFileName) = vbOK Then
            FNm = PathFileName
        End If
    Else
        If mnuOptionUseOldComDlg.Checked Then FNm = FileSaveOld Else FNm = FileSaveNew
    End If
    If Len(FNm) Then LblSFD.Caption = FNm
End Sub
Private Function FileSaveNew() As String
    With New SaveFileDialog
        '.Filter = MApp.FileExtFilter
        'Debug.Print .Filter
        '.ShowHelp = True
        If .ShowDialog = vbCancel Then Exit Function
        FileSaveNew = .FileName
    End With
End Function
Private Function FileSaveOld() As String
'Try: On Error GoTo Catch
'    With Me.CommonDialog
'        .Filter = MApp.FileExtFilter
'        .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
'        .flags = .flags Or FileOpenConstants.cdlOFNPathMustExist
'        .DefaultExt = ".htm"
'        .CancelError = True
'        .flags = .flags Or FileOpenConstants.cdlOFNReadOnly
'        .ShowOpen
'        FileSaveOld = .FileName
'    End With
'Catch:
'    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
'        MComDlgCtrl.MessCommonDlgError Err.Number
'    End If
End Function

'Private Sub mnuFilePrinter1_Click()
'    Dim PNm As String
'    If mnuOptionUseOldComDlg.Checked Then PNm = FilePrinterOld Else PNm = FilePrinterNew
'    If Len(PNm) = 0 Then Exit Sub
'    'If Len(PNm) Then MsgBox PNm
'
'    'MsgBox Printer.DeviceName
'    'MsgBox Printer.DriverName
'    'Dim pk As PaperKind: pk = Printer.PaperSize
'    'MsgBox pk & " = " & MPrinterPaper.PaperKind_ToStr(pk)
'
'End Sub

Private Sub mnuFilePrinter1_Click()
    Dim PDlg As New PrintDialog
    InitPrinterSettings PDlg
    If PDlg.ShowDialog() = vbCancel Then Exit Sub
    MsgBox ReportPrinterSettings(PDlg)
    Set Printer = SelectPrinter(PDlg.PrinterSettings_PrinterName)
End Sub

Private Sub mnuFilePrinter2_Click()
    
    MsgBox "Nope - does not work a.t.m."
    Exit Sub
    
    Dim PDlg As New PrintDialog
    InitPrinterSettings PDlg
    PDlg.UseEXDialog = True
    If PDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    MsgBox ReportPrinterSettings(PDlg)
    Set Printer = SelectPrinter(PDlg.PrinterSettings_PrinterName)
End Sub

Private Sub mnuFilePrinter3_Click()
    Dim PDlg As New PrintDialog
    InitPrinterSettings PDlg
    If PDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    MsgBox ReportPrinterSettings(PDlg)
    Set Printer = SelectPrinter(PDlg.PrinterSettings_PrinterName)
End Sub

Sub InitPrinterSettings(aPrintDlg As PrintDialog)
    With aPrintDlg
        .AllowCurrentPage = True
        .AllowPrintToFile = True
        .AllowSelection = True
        .AllowSomePages = True
        .ShowHelp = True
        .ShowNetwork = True
        .PrinterSettings_FromPage = 5
        .PrinterSettings_ToPage = 20
        .PrinterSettings_MinimumPage = 1
        .PrinterSettings_MaximumPage = 100
        .PrinterSettings_Copies = 14
    End With
End Sub

Function ReportPrinterSettings(aPrintDlg As PrintDialog) As String
    Dim s As String
    With aPrintDlg
        s = s & "PrinterSettings.PrinterName       : " & .PrinterSettings_PrinterName & vbCrLf
        s = s & "PrinterSettings.PrinterDriverName : " & .PrinterSettings_PrinterDriverName & vbCrLf
        s = s & "PrinterSettings.PrinterOutputName : " & .PrinterSettings_PrinterOutputName & vbCrLf
        s = s & "PrinterSettings.PrinterDefaultName: " & .PrinterSettings_PrinterDefaultName & vbCrLf
        s = s & "PrinterSettings.IsDefaultPrinter  : " & .PrinterSettings_IsDefaultPrinter & vbCrLf
        s = s & "PrinterSettings.Copies            : " & .PrinterSettings_Copies & vbCrLf
        s = s & "PrinterSettings.CanDuplex         : " & .PrinterSettings_CanDuplex & vbCrLf
        s = s & "PrinterSettings.LandscapeAngle    : " & .PrinterSettings_LandscapeAngle & vbCrLf
        s = s & "PrinterSettings.MaximumCopies     : " & .PrinterSettings_MaximumCopies & vbCrLf
        s = s & "PrinterSettings.MinimumPage       : " & .PrinterSettings_MinimumPage & vbCrLf
        s = s & "PrinterSettings.MaximumPage       : " & .PrinterSettings_MaximumPage & vbCrLf
        s = s & "PrinterSettings.FromPage          : " & .PrinterSettings_FromPage
        s = s & "PrinterSettings.ToPage            : " & .PrinterSettings_ToPage
        s = s & "PrinterSettings.PrintToFile       : " & .PrinterSettings_PrintToFile & vbCrLf
        s = s & "PrinterSettings.PrintFileName     : " & .PrinterSettings_PrintFileName & vbCrLf
        s = s & "PrinterSettings.SupportsColor     : " & .PrinterSettings_SupportsColor & vbCrLf
        's = s & "ShowHelp                          : " & .ShowHelp & vbCrLf
        s = s & "PrintToFile                       : " & .PrintToFile & vbCrLf
        
    End With
    ReportPrinterSettings = s
End Function
'
'Private Function FilePrinterOld() As String
''Try: On Error GoTo Catch
''    With Me.CommonDialog
''        .CancelError = True
''        .ShowPrinter
''        FilePrinterOld = Printer.DeviceName
''    End With
''Catch:
''    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
''        MComDlgCtrl.MessCommonDlgError Err.Number
''    End If
'End Function

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
'    ColorChooseOld = -1
'Try: On Error GoTo Catch
'    With CommonDialog
'        .Color = LblCD.BackColor
'        .CancelError = True
'        .ShowColor
'        ColorChooseOld = .Color
'    End With
'Catch:
'    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
'        MComDlgCtrl.MessCommonDlgError Err.Number
'    End If
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
'Try: On Error GoTo Catch
'    Dim Font As StdFont
'    With CommonDialog
'        .CancelError = True
'        .Color = Color_inout
'        .FontName = Font_in.Name
'        .FontSize = Font_in.Size
'        .FontBold = Font_in.Bold
'        .FontItalic = Font_in.Italic
'        .FontUnderline = Font_in.Underline
'        .FontStrikethru = Font_in.Strikethrough
'        .ShowFont
'    End With
'Catch:
'    Set FontDialogOld = Font_in
'    If Not Err.Number = MSComDlg.ErrorConstants.cdlCancel Then
'        MComDlgCtrl.MessCommonDlgError Err.Number
'    End If
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
    MsgBox "Nope - I am sorry - does not work anymore"
    'mnuOptionUseOldComDlg.Checked = Not mnuOptionUseOldComDlg.Checked
    'Dim bUseOldComDlg As Boolean: bUseOldComDlg = mnuOptionUseOldComDlg.Checked
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
                              .flags = .flags Or BIF_RETURNONLYFSDIRS
        Case CSIDL_NETWORK
                              .flags = 0 'set to 0 before!
                              .flags = .flags Or BIF_BROWSEFORCOMPUTER
        Case CSIDL_PRINTERS
                              .flags = .flags Or BIF_BROWSEFORPRINTER
        Case SpecialFolder_Personal
                              '.Flags = .Flags Or BIF_DONTGOBELOWDOMAIN
                              .flags = 0 'set to 0 before!
                              .flags = .flags Or BIF_RETURNFSANCESTORS
        End Select
        
        If ChkShowEditBox.Value = vbChecked Then
            .flags = .flags Or BIF_EDITBOX
        End If
        If Me.ChkShowNewFolderButton = vbUnchecked Then
            .flags = .flags Or BIF_DONTSHOWNEWFOLDERBUTTON
        End If
        'maximal 3 Zeilen Beschreibungstext
        .Description = "Hier sollte ein Hinweis stehen f�r den Benutzer was er hier tun soll. " & _
                       "In maximal 3 Zeilen erkl�rt. 12345 67890 12345 67890 12345 67890 " & _
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

'PrintDialog:
'https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/printdialog-component-windows-forms?view=netframeworkdesktop-4.8
'https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.printdialog?view=windowsdesktop-8.0
'
'PrinterSettings:
'
'https://learn.microsoft.com/en-us/dotnet/api/system.drawing.printing.printersettings?view=windowsdesktop-8.0
'
'PrintDocument:
'https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/printdocument-component-windows-forms?view=netframeworkdesktop-4.8
'https://learn.microsoft.com/en-us/dotnet/api/system.drawing.printing.printdocument?view=windowsdesktop-8.0
'
'PrintPreviewControl:
'https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/printpreviewcontrol-control-windows-forms?view=netframeworkdesktop-4.8
'https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.printpreviewcontrol?view=windowsdesktop-8.0
'
'PrintPreviewDialog:
'https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/printpreviewdialog-control-windows-forms?view=netframeworkdesktop-4.8
'https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.printpreviewdialog?view=windowsdesktop-8.0
'
'Private Sub ButtonPrint_Click(sender As Object, e as EventArgs) Handles ButtonPrint.Click
'
'    Printdialog1.Document = PrintDocument1
'
'    Printdialog1.PrinterSettings = PrintDocument1.PrinterSettings
'
'    Printdialog1.AllowSomePages = True
'
'    If Printdialog1.ShowDialog = DialogResult.OK Then
'
'        PrintDocument1.PrinterSettings = Printdialog1.PrinterSettings
'        PrintDocument1.Print()
'
'    End If
'
'    'what about
'    '* PrintPreviewDialog
'    '* PrintPreviewControl
'
'PrintDialog.PrintSettings.ToString
'[PrinterSettings
'    HP LaserJet CP 1025nw
'    Copies = 1
'    Collate = False
'    Duplex = Simplex
'    FromPage = 0
'    LandscapeAngle = 270
'    MaximumCopies = 999
'    OutputPort =
'    ToPage = 0
']
'
'End Sub
'   Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
'       Dim s1 As String, s2 As String, s3 As String
'       MessageBox.Show (PrintDialog1.ToString) 'System.Windows.Forms.PrintDialog
'       s1 = PrintDialog1.PrinterSettings.ToString
'       MessageBox.Show (s1)
'       '[PrinterSettings
'       '    HP LaserJet CP 1025nw
'       '    Copies = 1
'       '    Collate = False
'       '    Duplex = Simplex
'       '    FromPage = 0
'       '    LandscapeAngle = 270
'       '    MaximumCopies = 999
'       '    OutputPort =
'       '    ToPage = 0
'       ']
'       'MessageBox.Show(PrintDialog1.Document.ToString()) 'NullReferenceException
'       MessageBox.Show (PrintDocument1.ToString) '[PrintDocument document]
'       s2 = PrintDocument1.PrinterSettings.ToString
'       MessageBox.Show (s2)
'       If s1 = s2 Then MessageBox.Show ("OK s1 = s2")
'       PrintDialog1.Document = PrintDocument1
'       MessageBox.Show (PrintDialog1.Document.ToString) '[PrintDocument document]
'
'
'
'       'PrintDocument1.PrinterSettings.Copies = 14
'       PrintDialog1.AllowCurrentPage = True
'       PrintDialog1.AllowPrintToFile = True
'       PrintDialog1.AllowSelection = True
'       PrintDialog1.AllowSomePages = True
'
'       'PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
'       'PrintDialog1.UseEXDialog = True
'
'       If PrintDialog1.ShowDialog() = DialogResult.Cancel Then Return
'
'       PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
'
'       MessageBox.Show (PrintDocument1.PrinterSettings.Copies.ToString)
'       'PrintDocument1.Print()
'       'PrintPreviewDialog1.ShowDialog()
'   End Sub


Private Sub mnuOptionUseUndocumShell32_Click()
    mnuOptionUseUndocumShell32.Checked = Not mnuOptionUseUndocumShell32.Checked
End Sub
