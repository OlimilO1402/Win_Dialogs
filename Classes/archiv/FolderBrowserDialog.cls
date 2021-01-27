VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderBrowserDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class FolderBrowserDialog
'          Inherits System.Windows.Forms.CommonDialog
'     Member von: System.Windows.Forms
'
'Zusammenfassung:
' Stellt ein Standarddialogfeld dar, in dem Benutzer einen Ordner auswählen können.
' Die Klasse kann nicht vererbt werden.
Option Explicit
'Implements CommonDialog
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (pBrowseInfo As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function lstrcatA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As String) As Long
Private Type BrowseInfo
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
'ulFlags:
Private Const BIF_RETURNONLYFSDIRS   As Long = &H1 ': Gestattet nur Dateisystemordner als Auswahl.
Private Const BIF_DONTGOBELOWDOMAIN  As Long = &H2 ': Der Dialog zeigt keine Netzwerkordner unterhalb der aktuellen Domain.
Private Const BIF_STATUSTEXT         As Long = &H4 ': Der Dialog enthält eine Statuszeile. Die Rückruffunktion kann die Statuszeile ausfüllen.
Private Const BIF_RETURNFSANCESTORS  As Long = &H8 ': Gestattet nur Dateisystemobjekte als Auswahl
Private Const BIF_EDITBOX            As Long = &H10
Private Const BIF_VALIDATE           As Long = &H20
Private Const BIF_NEWDIALOGSTYLE     As Long = &H40
Private Const BIF_USENEWUI           As Long = &H40
Private Const BIF_BROWSEINCLUDEURLS  As Long = &H80
Private Const BIF_BROWSEFORCOMPUTER  As Long = &H1000 ': Als Auswahl sind nur Computer erlaubt. Wenn der Anwender andere Objekte, also Ordner oder Laufwerke markiert, kann der OK-Button nicht ausgewählt werden.
Private Const BIF_BROWSEFORPRINTER   As Long = &H2000 ': Gestattet nur Drucker als Auswahl.
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000 ': Der Dialog zeigt neben Computern, Laufwerken und Ordnern auch Dateien an.
Private Const BIF_SHAREABLE          As Long = &H8000

Private mFlags As Long
Private mDescription As String
Private mRootFolder As Long
Private mSelectedPath As String
Private mTag As Variant
Public Event HelpRequest() 'ByVal sender As Object, ByVal e As System.EventArgs)
Public Enum Environment_SpecialFolder
  Desktop = 0
  Programs = 2
  Personal = 5
  Favorites = 6
  Startup = 7
  Recent = 8
  SendTo = 9
  StartMenu = 11
  MyMusic = 13
  DesktopDirectory = 16
  MyComputer = 17
  Templates = 21
  ApplicationData = 26
  LocalApplicationData = 28
  InternetCache = 32
  Cookies = 33
  History = 34
  CommonApplicationData = 35
  System = 37
  ProgramFiles = 38
  MyPictures = 39
  CommonProgramFiles = 43
End Enum
Private Sub Class_Initialize()
  Call Reset
End Sub
'Public Sub New() 'wird nicht gebraucht

'Public Property Description() As String
Public Property Get Description() As String
' Ruft den beschreibenden Text ab, der im Dialogfeld über dem
' Strukturansichts-Steuerelement angezeigt wird, oder legt diesen fest.
  Description = mDescription
End Property
Public Property Let Description(StrVal As String)
  mDescription = StrVal
End Property

'Public Overrides Sub Reset()
Public Sub Reset()
  'mFlags = mFlags Or BIF_RETURNONLYFSDIRS
  'mFlags = mFlags Or BIF_DONTGOBELOWDOMAIN
  'mFlags = mFlags Or BIF_STATUSTEXT
  'mFlags = mFlags Or BIF_RETURNFSANCESTORS
  
  'mFlags = mFlags Or BIF_EDITBOX
  'mFlags = mFlags Or BIF_VALIDATE
  'mFlags = mFlags Or BIF_NEWDIALOGSTYLE
  mFlags = mFlags Or BIF_USENEWUI
  'mFlags = mFlags Or BIF_BROWSEINCLUDEURLS
  
  'mFlags = mFlags Or BIF_BROWSEFORCOMPUTER
  'mFlags = mFlags Or BIF_BROWSEFORPRINTER
  'mFlags = mFlags Or BIF_BROWSEINCLUDEFILES
  'mFlags = mFlags Or BIF_SHAREABLE
End Sub

'Public Property RootFolder() As System.Environment.SpecialFolder
Public Property Get RootFolder() As Environment_SpecialFolder 'System.Environment.SpecialFolder
' Ruft den Stammordner ab, von dem aus eine Suche gestartet wird, oder legt diesen fest.
  RootFolder = mRootFolder
End Property
Public Property Let RootFolder(LngVal As Environment_SpecialFolder) 'System.Environment.SpecialFolder
  mRootFolder = LngVal
End Property

'Public Property SelectedPath() As String
Public Property Get SelectedPath() As String
' Ruft den von den Benutzern ausgewählten Pfad ab oder legt diesen fest.
  SelectedPath = mSelectedPath
End Property
Public Property Let SelectedPath(StrVal As String)
  mSelectedPath = StrVal
End Property

'Public Property ShowNewFolderButton() As Boolean
Public Property Get ShowNewFolderButton() As Boolean
' Ruft den Wert ab, der angibt, ob die Schaltfläche New Folder im Dialogfeld für die
' Ordnersuche angezeigt wird.

End Property
Public Property Let ShowNewFolderButton(BolVal As Boolean)

End Property

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
  ToString = "MBO.Windows.Forms.FolderBrowserDialog"
End Function

Public Function ShowDialog(Optional Frm As Variant) As DialogResult
Dim BI As BrowseInfo
Dim hhwndOwner As Long
Dim n As Integer
Dim IDList As Long
Dim Buffer As String
  If IsMissing(Frm) Then
    hhwndOwner = GetActiveWindow
  Else
    hhwndOwner = Frm.hwnd
  End If
  
TryE: On Error GoTo CatchE
  
  With BI
    .hwndOwner = hhwndOwner
    .pidlRoot = mRootFolder
    '.pszDisplayName =
    'Titel des Dialoges als nullterminierter String
    .lpszTitle = mDescription & vbNullChar
    .ulFlags = mFlags
    '.lpfn = ""
    '.lParam = 0&
    '.iImage = 0&
  End With
  
  'Anzeigen des Dialogs und Übergabe an eine IID-Liste
  IDList = SHBrowseForFolder(BI)
  If IDList > 0 Then
    Buffer = String$(1024, vbNullChar)
    Call SHGetPathFromIDList(IDList, Buffer)
    mSelectedPath = Left(Buffer, lstrlenA(Buffer))
    ShowDialog = DialogResult_OK
  Else
    ShowDialog = DialogResult_Cancel
  End If
  Exit Function
CatchE:
  ShowDialog = DialogResult_Abort
End Function

'NN    FBD.Container.
'NN    FBD.CreateObjRef()
'OK    FBD.Description()
'NN    FBD.Dispose()
'      FBD.Equals
'      FBD.GetHashCode
'NN    FBD.GetLifetimeService()
'NN    FBD.GetType()
'NN    FBD.InitializeLifetimeService()
'      FBD.ReferenceEquals()
'OK    FBD.Reset()
'OK    FBD.RootFolder()
'OK    FBD.SelectedPath()
'OK    FBD.ShowDialog()
'--    FBD.ShowNewFolderButton()
'NA    FBD.Site()
'      FBD.Tag
'OK    FBD.ToString()

