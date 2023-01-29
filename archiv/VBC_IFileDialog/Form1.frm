VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Pick Folder Dialog"
      Height          =   555
      Left            =   7170
      TabIndex        =   4
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Custom Dialog"
      Height          =   555
      Left            =   5400
      TabIndex        =   3
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton Command3 
      Caption         =   "File Save Dialog"
      Height          =   555
      Left            =   3630
      TabIndex        =   2
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File Open Dialog"
      Height          =   555
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple Dialog"
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' IFileDialog und IFileDialog2 ist die Basis für Öffnen- oder Speichern-Dialoge.
' IFileDialog und IFileDialog2 können nur für SingleSelect verwendet werden. -> GetResult -> IShellItem
' Für MultiSelect ist das Interface IFileOpenDialog zuständig. -> GetResults -> IShellItemArray -> IEnumShellItems -> IShellItem
' Alle Dialoge können über IFileDialogCustomize angepasst werden.
' Bei allen Dialogen können die Events ausgewertet werden. -> IFileDialogEvents
' Microsoft empfiehlt sogar anstelle des BrowseForFolder-Dialogs diesen Dialog mit Option FOS_PICKFOLDERS (SingelSelect) zu verwenden.
' Für weitere Möglichkeiten mit den Dialogen, bitte in die entsprechenden MSDN-Artikel zu den Interfaces schauen.

Private Const IID_IFileDialogCustomize As String = "{e6fdd21a-163f-4975-9c8c-a69f1ba37034}"

' Klasse für die DialogEvents
Private WithEvents cIFileDialogEvents As clsIFileDialogEvents
Attribute cIFileDialogEvents.VB_VarHelpID = -1

Private Sub Command1_Click()

    Dim lCookie As Long
    Dim pIShellItemRet As Long
    Dim IShellItemRet As clsIShellItem
    ' Dim IFileDialog As clsIFileDialog
    Dim IFileDialog As clsIFileDialog2
    Dim eDlgOptions As FILEOPENDIALOGOPTIONS
    Dim tDlgFileFilter(1) As COMDLG_FILTERSPEC

    ' Set IFileDialog = New clsIFileDialog
    Set IFileDialog = New clsIFileDialog2
    
    ' Open oder Save-Dialog
    Call IFileDialog.InitAs(FileOpenDialog)

    ' FileFilter
    tDlgFileFilter(0).pszName = "Image Files"
    tDlgFileFilter(0).pszSpec = "*.jpg;*.gif;*.bmp"
    tDlgFileFilter(1).pszName = "All Files"
    tDlgFileFilter(1).pszSpec = "*.*"
    Call IFileDialog.SetFileTypes(tDlgFileFilter)

    ' FileFilterIndex
    Call IFileDialog.SetFileTypeIndex(2)
    ' Debug.Print "GetFileTypeIndex = " & CStr(IFileDialog.GetFileTypeIndex)

    ' FileName
    Call IFileDialog.SetFileName("test.txt")
    ' Debug.Print "GetFileName = " & IFileDialog.GetFileName

    ' Options
    eDlgOptions = IFileDialog.GetOptions
    eDlgOptions = eDlgOptions Or FOS_FILEMUSTEXIST Or FOS_FORCESHOWHIDDEN
    Call IFileDialog.SetOptions(eDlgOptions)

    ' DialogTitel
    Call IFileDialog.SetTitle("Simple File Open")

    ' OkButtonLabel
    Call IFileDialog.SetOkButtonLabel("Click here")

    ' CancelButtonLabel (nur in IFileDialog2)
    Call IFileDialog.SetCancelButtonLabel("Close me")

    ' FileNameLabel
    Call IFileDialog.SetFileNameLabel("FileName here ->")
    
    ' Optional!!! Ein IFileDialogEvents Interface erstellen
    Set cIFileDialogEvents = New clsIFileDialogEvents
    
    ' Optional!!! Events an das Interface IFileDialogEvents leiten
    lCookie = IFileDialog.Advise(cIFileDialogEvents.GetIPtr)

    ' Show Dialog
    If IFileDialog.Show(Me.hwnd) = True Then

        ' Return IShellItem
        pIShellItemRet = IFileDialog.GetResult

        ' gibt es einen Pointer auf ein IShellItem Interface
        If pIShellItemRet <> 0 Then

            ' Klasse für IShellItem initialisieren
            Set IShellItemRet = New clsIShellItem
            Call IShellItemRet.Init(pIShellItemRet)

            ' FILESYSPATH vom IShellItem auslesen
            Debug.Print IShellItemRet.GetDisplayName(SIGDN_FILESYSPATH)

            ' Aufräumen
            Set IShellItemRet = Nothing

        End If

    End If

    ' Optional!!! nur wenn über IFileDialog.Advise ein
    ' IFileDialogEvents Interface übergeben wurde.
    If lCookie <> 0 Then
    
        ' Aufräumen
        Call IFileDialog.Unadvise(lCookie)
        
    End If
    
    ' Aufräumen
    Set cIFileDialogEvents = Nothing
    
    ' Aufräumen
    Set IFileDialog = Nothing

End Sub

Private Sub Command2_Click()

    Dim lCookie As Long
    Dim pIShellItemArrayRet As Long
    Dim pIShellItem() As Long
    Dim pIEnumShellItems As Long
    Dim eDlgOptions As FILEOPENDIALOGOPTIONS
    Dim IShellItem As clsIShellItem
    Dim IFileOpenDialog As clsIFileOpenDialog
    Dim IShellItemArray As clsIShellItemArray
    Dim IEnumShellItems As clsIEnumShellItems
    Dim tDlgFileFilter(0) As COMDLG_FILTERSPEC

    Set IFileOpenDialog = New clsIFileOpenDialog

    tDlgFileFilter(0).pszName = "All Files"
    tDlgFileFilter(0).pszSpec = "*.*"
    Call IFileOpenDialog.SetFileTypes(tDlgFileFilter)

    eDlgOptions = IFileOpenDialog.GetOptions
    eDlgOptions = eDlgOptions Or FOS_ALLOWMULTISELECT ' MultiSelect Demo!!!
    Call IFileOpenDialog.SetOptions(eDlgOptions)

    Set cIFileDialogEvents = New clsIFileDialogEvents
    lCookie = IFileOpenDialog.Advise(cIFileDialogEvents.GetIPtr)

    If IFileOpenDialog.Show(Me.hwnd) = True Then

        ' bei MultiSelect: IFileOpenDialog.GetResults -> IShellItemArray
        pIShellItemArrayRet = IFileOpenDialog.GetResults

        ' ist ein Pointer auf ein IShellItemArray Interface vorhanden
        If pIShellItemArrayRet <> 0 Then

            ' Klasse für IShellItemArray initialisieren
            Set IShellItemArray = New clsIShellItemArray
            Call IShellItemArray.Init(pIShellItemArrayRet)

            ' IShellItemArray.EnumItems -> IEnumShellItems
            pIEnumShellItems = IShellItemArray.EnumItems

            ' ist ein Pointer auf ein IEnumShellItems Interface vorhanden
            If pIEnumShellItems <> 0 Then

                ' Klasse für IEnumShellItems initialisieren
                Set IEnumShellItems = New clsIEnumShellItems
                Call IEnumShellItems.Init(pIEnumShellItems)

                ' alle Items in IEnumShellItems durchlaufen
                Do While IEnumShellItems.Fetch(1, pIShellItem, 0) = True

                    ' ist ein Pointer auf ein IShellItem vorhanden
                    If pIShellItem(0) <> 0 Then

                        ' Klasse für IShellItem initialisieren
                        Set IShellItem = New clsIShellItem
                        Call IShellItem.Init(pIShellItem(0))

                        ' FILESYSPATH von IShellItem auslesen
                        Debug.Print IShellItem.GetDisplayName(SIGDN_FILESYSPATH)

                        ' Aufräumen
                        Set IShellItem = Nothing

                    End If

                Loop

                ' Aufräumen
                Set IEnumShellItems = Nothing

            End If

            ' Aufräumen
            Set IShellItemArray = Nothing

        End If

    End If

    If lCookie <> 0 Then
    
        ' Aufräumen
        Call IFileOpenDialog.Unadvise(lCookie)
            
    End If
    
    ' Aufräumen
    Set cIFileDialogEvents = Nothing

    ' Aufräumen
    Set IFileOpenDialog = Nothing

End Sub

Private Sub Command3_Click()

    Dim lCookie As Long
    Dim pIShellItemRet As Long
    Dim IShellItemRet As clsIShellItem
    Dim tDlgFileFilter(1) As COMDLG_FILTERSPEC
    Dim IFileSaveDialog As clsIFileSaveDialog

    Set IFileSaveDialog = New clsIFileSaveDialog

    tDlgFileFilter(0).pszName = "Image Files"
    tDlgFileFilter(0).pszSpec = "*.jpg;*.gif;*.bmp"
    tDlgFileFilter(1).pszName = "All Files"
    tDlgFileFilter(1).pszSpec = "*.*"
    Call IFileSaveDialog.SetFileTypes(tDlgFileFilter)
    Call IFileSaveDialog.SetDefaultExtension(".jpg")
    Call IFileSaveDialog.SetFileName("NewFile.jpg")

    Set cIFileDialogEvents = New clsIFileDialogEvents
    lCookie = IFileSaveDialog.Advise(cIFileDialogEvents.GetIPtr)

    If IFileSaveDialog.Show(Me.hwnd) = True Then

        pIShellItemRet = IFileSaveDialog.GetResult

        If pIShellItemRet <> 0 Then

            Set IShellItemRet = New clsIShellItem
            Call IShellItemRet.Init(pIShellItemRet)

            Debug.Print IShellItemRet.GetDisplayName(SIGDN_FILESYSPATH)

            Set IShellItemRet = Nothing

        End If

    End If

    If lCookie <> 0 Then
    
        Call IFileSaveDialog.Unadvise(lCookie)
        
    End If

    Set cIFileDialogEvents = Nothing

    Set IFileSaveDialog = Nothing

End Sub

Private Sub Command4_Click()

    Dim lCookie As Long
    Dim pIShellItemRet As Long
    Dim pIFileDialogCustomize As Long
    Dim IShellItemRet As clsIShellItem
    Dim IFileDialogCustomize As clsIFileDialogCustomize
    Dim IFileDialog As clsIFileDialog
    Dim eDlgOptions As FILEOPENDIALOGOPTIONS
    Dim tDlgFileFilter(1) As COMDLG_FILTERSPEC

    Set IFileDialog = New clsIFileDialog
    Call IFileDialog.InitAs(FileOpenDialog)

    ' FileFilter
    tDlgFileFilter(0).pszName = "Text Files"
    tDlgFileFilter(0).pszSpec = "*.txt"
    tDlgFileFilter(1).pszName = "All Files"
    tDlgFileFilter(1).pszSpec = "*.*"
    Call IFileDialog.SetFileTypes(tDlgFileFilter)

    ' FileFilterIndex
    Call IFileDialog.SetFileTypeIndex(1)

    ' Options
    eDlgOptions = IFileDialog.GetOptions
    eDlgOptions = eDlgOptions Or FOS_FILEMUSTEXIST Or FOS_FORCESHOWHIDDEN Or FOS_FORCEPREVIEWPANEON
    Call IFileDialog.SetOptions(eDlgOptions)

    ' ein IFileDialogCustomize Interface erstellen
    pIFileDialogCustomize = IFileDialog.QueryInterface(IID_IFileDialogCustomize)

    ' ist ein IFileDialogCustomize Interface vorhanden
    If pIFileDialogCustomize <> 0 Then

        ' Klasse für IFileDialogCustomize initialisieren
        Set IFileDialogCustomize = New clsIFileDialogCustomize
        Call IFileDialogCustomize.Init(pIFileDialogCustomize)
        
        ' IFileDialog modifizieren (Editor Style)
        Call IFileDialogCustomize.StartVisualGroup(1000, "Codierung:")
        Call IFileDialogCustomize.AddComboBox(2000)
        Call IFileDialogCustomize.AddControlItem(2000, 2001, "ANSI")
        Call IFileDialogCustomize.AddControlItem(2000, 2002, "Unicode")
        Call IFileDialogCustomize.AddControlItem(2000, 2003, "Unicode Big Endian")
        Call IFileDialogCustomize.AddControlItem(2000, 2004, "UTF-8")
        Call IFileDialogCustomize.SetSelectedControlItem(2000, 2001)
        Call IFileDialogCustomize.EndVisualGroup
        
        Set cIFileDialogEvents = New clsIFileDialogEvents
        lCookie = IFileDialog.Advise(cIFileDialogEvents.GetIPtr)
        
        If IFileDialog.Show(Me.hwnd) = True Then

            pIShellItemRet = IFileDialog.GetResult

            If pIShellItemRet <> 0 Then

                Set IShellItemRet = New clsIShellItem

                Call IShellItemRet.Init(pIShellItemRet)

                Debug.Print IShellItemRet.GetDisplayName(SIGDN_FILESYSPATH)

                Set IShellItemRet = Nothing

            End If

        End If

        If lCookie <> 0 Then
    
            Call IFileDialog.Unadvise(lCookie)
      
        End If
        
        Set cIFileDialogEvents = Nothing

        Set IFileDialogCustomize = Nothing

        Set IFileDialog = Nothing

    End If

End Sub

Private Sub Command5_Click()

    Dim pIShellItemRet As Long
    Dim IShellItemRet As clsIShellItem
    Dim IFileDialog As clsIFileDialog
    Dim eDlgOptions As FILEOPENDIALOGOPTIONS
    Dim tDlgFileFilter(1) As COMDLG_FILTERSPEC

    Set IFileDialog = New clsIFileDialog
    Call IFileDialog.InitAs(FileOpenDialog)
    
    eDlgOptions = IFileDialog.GetOptions
    eDlgOptions = eDlgOptions Or FOS_PICKFOLDERS ' PICKFOLDERS Demo!!!
    Call IFileDialog.SetOptions(eDlgOptions)

    Call IFileDialog.SetTitle("PICKFOLDERS Dialog")

    If IFileDialog.Show(Me.hwnd) = True Then

        pIShellItemRet = IFileDialog.GetResult

        If pIShellItemRet <> 0 Then

            Set IShellItemRet = New clsIShellItem
            Call IShellItemRet.Init(pIShellItemRet)

            Debug.Print IShellItemRet.GetDisplayName(SIGDN_FILESYSPATH)

            Set IShellItemRet = Nothing

        End If

    End If

    Set IFileDialog = Nothing

End Sub

' ----==== IFileDialogEvents ====----
Private Sub cIFileDialogEvents_OnFileOk(ByVal pfd As Long)

    Debug.Print "OnFileOk"
    
End Sub

Private Sub cIFileDialogEvents_OnFolderChanging(ByVal pfd As Long, ByVal _
    psiFolder As Long)

    Debug.Print "OnFolderChanging"

End Sub

Private Sub cIFileDialogEvents_OnFolderChange(ByVal pfd As Long)

    Debug.Print "OnFolderChange"

End Sub

Private Sub cIFileDialogEvents_OnSelectionChange(ByVal pfd As Long)

    Debug.Print "OnSelectionChange"

End Sub

Private Sub cIFileDialogEvents_OnShareViolation(ByVal pfd As Long, ByVal psi _
    As Long)

    Debug.Print "OnShareViolation"

End Sub

Private Sub cIFileDialogEvents_OnTypeChange(ByVal pfd As Long)

    Debug.Print "OnTypeChange"

End Sub

Private Sub cIFileDialogEvents_OnOverwrite(ByVal pfd As Long, ByVal psi As _
    Long)

    Debug.Print "OnOverwrite"

End Sub

' ----==== IFileDialogControlEvents ====----
Private Sub cIFileDialogEvents_OnItemSelected(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal dwIDItem As Long)

    Debug.Print "OnItemSelected", dwIDCtl, dwIDItem

End Sub

Private Sub cIFileDialogEvents_OnButtonClicked(ByVal pfdc As Long, ByVal dwIDCtl As Long)

    Debug.Print "OnButtonClicked", dwIDCtl

End Sub

Private Sub cIFileDialogEvents_OnCheckButtonToggled(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal bChecked As Boolean)

    Debug.Print "OnCheckButtonToggled", dwIDCtl, bChecked

End Sub

Private Sub cIFileDialogEvents_OnControlActivating(ByVal pfdc As Long, ByVal dwIDCtl As Long)

    Debug.Print "OnControlActivating", dwIDCtl

End Sub

