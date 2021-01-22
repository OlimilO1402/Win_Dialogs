VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "lokaler Ordner"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Drucker"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Computer"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ordner"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton BtnFolderBrowser 
      Caption         =   "FolderBrowser"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   5055
   End
   Begin VB.ComboBox CmbSpecialFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3360
      Width           =   5055
   End
   Begin VB.CheckBox ChkShowEditBox 
      Caption         =   "ShowEditBox"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.CheckBox ChkShowNewFolderButton 
      Caption         =   "ShowNewFolderButton"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   2880
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.TextBox TxtSelectedPath 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   5055
   End
   Begin VB.CheckBox ChkSelectedPath 
      Caption         =   "Verwende falls möglich obige TextBox als voreingestellten Pfad"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Value           =   1  'Aktiviert
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Der FolderBrowserDialog als spezieller Dialog für das Suchen von ..."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4320
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
      Caption         =   "&Datei"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Öffnen..."
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Speichern &unter..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Be&enden"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuEditColorChoose 
         Caption         =   "&Farbe auswählen..."
      End
      Begin VB.Menu mnuEditFontChoose 
         Caption         =   "Schrift &auswählen..."
      End
      Begin VB.Menu mnuEditPathChoose 
         Caption         =   "&Pfad auswählen..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OFD As New OpenFileDialog
Private SFD As New SaveFileDialog
Private FBD As New FolderBrowserDialog
Private CD  As New ColorDialog
Private FD  As New FontDialog

Private Sub Form_Load()
    PrepareSpecialFolder
End Sub

Private Sub mnuFileOpen_Click()
    OFD.Filter = "TextDatei (*.txt)|*.txt|html-Datei (*.htm, *.html)|*.htm*|Alle Dateien (*.*)|*.*"
    OFD.CheckFileExists = False
    OFD.CheckPathExists = False
    OFD.DefaultExt = ".htm"
    OFD.ShowReadOnly = True
    OFD.AddExtension = False
    If OFD.ShowDialog() = DialogResult.DialogResult_OK Then
        LblOFD.Caption = OFD.FileName
    End If
End Sub
Private Sub mnuFileSaveAs_Click()
    SFD.Filter = "TextDatei (*.txt)|*.txt|html-Datei (*.htm, *.html)|*.htm*|Alle Dateien (*.*)|*.*"
    If SFD.ShowDialog = DialogResult.DialogResult_OK Then
        LblSFD.Caption = SFD.FileName
    End If
End Sub
'--------------------------------------------------
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
'==================================================
Private Sub mnuEditPathChoose_Click()
    FBD.Description = "Wählen Sie bitte einen Ordner aus:"
    FBD.ShowNewFolderButton = True
    If FBD.ShowDialog = DialogResult_OK Then
        LblFBD.Caption = FBD.SelectedPath
    End If
End Sub

Private Sub mnuEditColorChoose_Click()
    CD.Color = LblCD.BackColor
    CD.SolidColorOnly = True
    If CD.ShowDialog = DialogResult_OK Then
        LblCD.BackColor = CD.Color
    End If
End Sub

Private Sub mnuEditFontChoose_Click()
    Set FD.Font = LblFD.Font
    If FD.ShowDialog = DialogResult_OK Then
        Set LblFD.Font = FD.Font
    End If
End Sub


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
                              .flags = 0 'vorher zu null setzen!
                              .flags = .flags Or BIF_BROWSEFORCOMPUTER
        Case CSIDL_PRINTERS
                              .flags = .flags Or BIF_BROWSEFORPRINTER
        Case SpecialFolder_Personal
                              '.Flags = .Flags Or BIF_DONTGOBELOWDOMAIN
                              .flags = 0
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
        If .ShowDialog = DialogResult_OK Then
            TxtSelectedPath.Text = .SelectedPath
        End If
    End With

End Sub



