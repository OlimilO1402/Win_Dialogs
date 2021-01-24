VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "locale Folders"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Printer"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Computer"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Folders"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton BtnFolderBrowser 
      Caption         =   "FolderBrowser"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   5055
   End
   Begin VB.ComboBox CmbSpecialFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3720
      Width           =   5055
   End
   Begin VB.CheckBox ChkShowEditBox 
      Caption         =   "ShowEditBox"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.CheckBox ChkShowNewFolderButton 
      Caption         =   "ShowNewFolderButton"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   3240
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.TextBox TxtSelectedPath 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   5055
   End
   Begin VB.CheckBox ChkSelectedPath 
      Caption         =   "Use the path above as the starting folder if possible"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Value           =   1  'Aktiviert
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":1782
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Use FolderBrowserDialog as special dialog fo searching ..."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4680
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CD  As New ColorDialog

Private Sub Form_Load()
    PrepareSpecialFolder
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
            Dim Pos As Long: Pos = InStrRev(FNm, "\")
            If Pos > 3 Then
                FNm = Left(FNm, Pos)
                LblFBD.Caption = FNm
            End If
        End If
    End With
End Sub

Private Sub mnuFileOpen_Click()
    With New OpenFileDialog
        .Filter = "TextDatei (*.txt)|*.txt|html-Datei (*.htm, *.html)|*.htm*|Alle Dateien (*.*)|*.*"
        .CheckFileExists = False
        .CheckPathExists = False
        .DefaultExt = ".htm"
        .ShowReadOnly = True
        .AddExtension = False
        If .ShowDialog = vbOK Then
            LblOFD.Caption = .FileName
        End If
    End With
End Sub
Private Sub mnuFileSaveAs_Click()
    With New SaveFileDialog
        .Filter = "TextDatei (*.txt)|*.txt|html-Datei (*.htm, *.html)|*.htm*|Alle Dateien (*.*)|*.*"
        If .ShowDialog = vbOK Then
            LblSFD.Caption = .FileName
        End If
    End With
End Sub
'--------------------------------------------------
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
'==================================================
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

Private Sub mnuEditColorChoose_Click()
    With CD
        .Color = LblCD.BackColor
        .SolidColorOnly = True
        If .ShowDialog = vbOK Then
            LblCD.BackColor = .Color
        End If
    End With
End Sub

Private Sub mnuEditFontChoose_Click()
    With New FontDialog
        .Font = LblFD.Font
        If .ShowDialog = vbOK Then
            Set LblFD.Font = .Font
        End If
    End With
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
        .Description = "Hier sollte ein Hinweis stehen für den Benutzer was er hier tun soll. " & _
                       "In maximal 3 Zeilen erklärt. 12345 67890 12345 67890 12345 67890 " & _
                       "12345 67890 12345 67890 12345 67890 12345!"
        If (ChkSelectedPath.Value = vbChecked) And (Len(TxtSelectedPath.Text) > 0) Then
            .SelectedPath = TxtSelectedPath.Text
        End If
        If .ShowDialog(Me) = vbOK Then
            TxtSelectedPath.Text = .SelectedPath
        End If
    End With

End Sub



