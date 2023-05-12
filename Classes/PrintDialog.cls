VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Enum PrinterConstants
'    cdlPDAllPages      = 0
'    cdlPDSelection     = 1
'    cdlPDPageNums      = 2
'    cdlPDNoSelection   = 4
'    cdlPDNoPageNums    = 8
'    cdlPDCollate       = 16 (&H10)
'    cdlPDPrintToFile   = 32 (&H20)
'    cdlPDPrintSetup    = 64 (&H40)
'    cdlPDNoWarning     = 128 (&H80)
'    cdlPDReturnDC      = 256 (&H100)
'    cdlPDReturnIC      = 512 (&H200)
'    cdlPDReturnDefault = 1024 (&H400)
'    cdlPDHelpButton    = 2048 (&H800)
'    cdlPDUseDevModeCopies   = 262144  (&H40000)
'    cdlPDDisablePrintToFile = 524288  (&H80000)
'    cdlPDHidePrintToFile    = 1048576 (&H100000)

Private Const PD_ALLPAGES                   As Long = &H0&     ' cdlPDAllPages    = 0     ' Die Standardkennzeichnung, die angibt, dass das Optionsfeld "Alle " anfänglich ausgewählt ist. Dieses Kennzeichen wird als Platzhalter verwendet, um anzugeben, dass die PD_PAGENUMS und PD_SELECTION Flags nicht angegeben werden.
Private Const PD_SELECTION                  As Long = &H1&     ' cdlPDSelection   = 1     ' Wenn dieses Kennzeichen festgelegt ist, wird das Optionsfeld "Auswahl " ausgewählt. Wenn weder PD_PAGENUMS noch PD_SELECTION festgelegt ist, wird das Optionsfeld "Alle " ausgewählt.
Private Const PD_PAGENUMS                   As Long = &H2&     ' cdlPDPageNums    = 2     ' Wenn dieses Kennzeichen festgelegt ist, wird das Optionsfeld "Seiten " ausgewählt. Wenn dieses Kennzeichen festgelegt ist, wenn die PrintDlg-Funktion zurückgegeben wird, geben die Elemente "nFromPage " und "nToPage " die vom Benutzer angegebenen Start- und Endseiten an.
Private Const PD_NOSELECTION                As Long = &H4&     ' cdlPDNoSelection = 4     ' Deaktiviert das Optionsfeld "Auswahl ".
Private Const PD_NOPAGENUMS                 As Long = &H8&     ' cdlPDNoPageNums  = 8     ' Deaktiviert das Optionsfeld "Seiten " und die zugehörigen Bearbeitungssteuerelemente. Außerdem wird das Kontrollkästchen "Zusammenzählen " im Dialogfeld angezeigt.
Private Const PD_COLLATE                    As Long = &H10&    ' cdlPDCollate     = 16 (&H10) ' Wenn dieses Kennzeichen festgelegt ist, ist das Kontrollkästchen "Zusammensetzen " aktiviert. Wenn dieses Kennzeichen festgelegt wird, wenn die PrintDlg-Funktion zurückgegeben wird, muss die Anwendung die Sortierung mehrerer Kopien simulieren. Weitere Informationen finden Sie in der Beschreibung des PD_USEDEVMODECOPIESANDCOLLATE Flags. Siehe PD_NOPAGENUMS.
Private Const PD_PRINTTOFILE                As Long = &H20&    ' cdlPDPrintToFile = 32 (&H20) ' Wenn dieses Kennzeichen festgelegt ist, ist das Kontrollkästchen "In Datei drucken " aktiviert. Wenn dieses Flag festgelegt wird, wenn die PrintDlg-Funktion zurückgegeben wird, enthält der offset, der durch das wOutputOffset-Element der DEVNAMES-Struktur angegeben ist, die Zeichenfolge "FILE:". Wenn Sie die StartDoc-Funktion aufrufen, um den Druckvorgang zu starten, geben Sie diese "FILE:"-Zeichenfolge im lpszOutput-Element der DOCINFO-Struktur an. Wenn Sie diese Zeichenfolge angeben, wird das Drucksubsystem zum Abfragen des Benutzers für den Namen der Ausgabedatei führen.
Private Const PD_PRINTSETUP                 As Long = &H40&    ' cdlPDPrintSetup  = 64 (&H40) ' Bewirkt, dass das System das Dialogfeld " Drucken einrichten " anstelle des Dialogfelds "Drucken " anzeigt.
Private Const PD_NOWARNING                  As Long = &H80&    ' cdlPDNoWarning  = 128 (&H80) ' Verhindert, dass die Warnmeldung angezeigt wird, wenn kein Standarddrucker vorhanden ist.
Private Const PD_RETURNDC                   As Long = &H100&   ' cdlPDReturnDC   = 256 (&H100) ' Bewirkt, dass PrintDlg einen Gerätekontext zurückgibt, der mit den Auswahlen übereinstimmen, die der Benutzer im Dialogfeld erstellt hat. Der Gerätekontext wird in hDC zurückgegeben.
Private Const PD_RETURNIC                   As Long = &H200&   ' cdlPDReturnIC   = 512 (&H200) ' Ähnlich wie bei der PD_RETURNDC-Kennzeichnung gibt dieses Flag anstelle eines Gerätekontexts einen Informationskontext zurück. Wenn weder PD_RETURNDC noch PD_RETURNIC angegeben wird, wird hDC bei der Ausgabe nicht definiert.
Private Const PD_RETURNDEFAULT              As Long = &H400&   ' cdlPDReturnDefault = 1024 (&H400) ' Wenn dieses Kennzeichen festgelegt ist, zeigt die PrintDlg-Funktion das Dialogfeld nicht an. Stattdessen werden die hDevNames - und hDevMode-Elemente so festgelegt, dass sie auf DEVMODE - und DEVNAMES-Strukturen verarbeitet werden, die für den Systemstandarddrucker initialisiert werden. Sowohl hDevNames als auch hDevMode müssen NULL sein, oder PrintDlg gibt einen Fehler zurück.
Private Const PD_SHOWHELP                   As Long = &H800&   ' cdlPDHelpButton    = 2048 (&H800) ' Bewirkt, dass das Dialogfeld die Hilfeschaltfläche anzeigt. Das hwndOwner-Element muss das Fenster angeben, um die registrierten HELPMSGSTRING-Nachrichten zu empfangen, die das Dialogfeld sendet, wenn der Benutzer auf die Hilfeschaltfläche klickt.
Private Const PD_ENABLEPRINTHOOK            As Long = &H1000&  '    Aktiviert die im Element "lpfnPrintHook " angegebene Hook-Prozedur. Dadurch wird die Hook-Prozedur für das Dialogfeld "Drucken " aktiviert.
Private Const PD_ENABLESETUPHOOK            As Long = &H2000&  '    Aktiviert die im Element "lpfnSetupHook " angegebene Hook-Prozedur. Dadurch wird die Hook-Prozedur für das Dialogfeld "Drucken einrichten " aktiviert.
Private Const PD_ENABLEPRINTTEMPLATE        As Long = &H4000&  '    Gibt an, dass die Elemente "hInstance " und "lpPrintTemplateName " einen Ersatz für die Standardvorlage " Drucken" angeben.
Private Const PD_ENABLESETUPTEMPLATE        As Long = &H8000&  '    Gibt an, dass die Elemente "hInstance " und "lpSetupTemplateName " einen Ersatz für die Standardvorlage " Drucken einrichten " angeben.
Private Const PD_ENABLEPRINTTEMPLATEHANDLE  As Long = &H10000  '    Gibt an, dass das hPrintTemplate-Element einen Datenblock identifiziert, der eine vordefinierte Dialogfeldvorlage enthält. Diese Vorlage ersetzt die Standardvorlage für das Dialogfeld "Drucken" . Das System ignoriert das lpPrintTemplateName-Element , wenn dieses Flag angegeben ist.
Private Const PD_ENABLESETUPTEMPLATEHANDLE  As Long = &H20000  '    Gibt an, dass das hSetupTemplate-Element einen Datenblock identifiziert, der eine vordefinierte Dialogfeldvorlage enthält. Diese Vorlage ersetzt die Standardvorlage für das Dialogfeld "Drucken einrichten ". Das System ignoriert das lpSetupTemplateName-Element , wenn dieses Flag angegeben ist.
Private Const PD_USEDEVMODECOPIES           As Long = &H40000  ' cdlPDUseDevModeCopies   = 262144  (&H40000) ' Identisch mit PD_USEDEVMODECOPIESANDCOLLATE.
Private Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000  ' cdlPDDisablePrintToFile = 524288  (&H80000) ' Dieses Flag gibt an, ob Ihre Anwendung mehrere Kopien und Sortierung unterstützt. Legen Sie dieses Kennzeichen für eingabe fest, um anzugeben, dass Ihre Anwendung mehrere Kopien und Sortierung nicht unterstützt. In diesem Fall gibt das nCopies-Element der PRINTDLG-Struktur immer 1 zurück, und PD_COLLATE wird nie im Flags-Element festgelegt.
'Wenn dieses Kennzeichen nicht festgelegt ist, ist die Anwendung für das Drucken und Sortieren mehrerer Kopien verantwortlich. In diesem Fall gibt das nCopies-Element der PRINTDLG-Struktur die Anzahl der Kopien an, die der Benutzer drucken möchte, und das PD_COLLATE Flag im Flags-Element gibt an, ob der Benutzer die Sortierung wünscht.
'Unabhängig davon, ob dieses Kennzeichen festgelegt ist, kann eine Anwendung anhand von nCopies und PD_COLLATE ermitteln, wie viele Kopien gerendert werden sollen und ob sie sortiert gedruckt werden sollen.
'Wenn dieses Kennzeichen festgelegt ist und der Druckertreiber mehrere Kopien nicht unterstützt, ist das Bearbeitungssteuerelement " Kopien " deaktiviert. Wenn dieses Kennzeichen festgelegt ist und der Druckertreiber die Sortierung nicht unterstützt, ist das Kontrollkästchen "Zusammensetzen " deaktiviert.
'Die Elemente dmCopies und dmCollate der DEVMODE-Struktur enthalten die Kopien und Sortierung von Informationen, die vom Druckertreiber verwendet werden. Wenn dieses Kennzeichen festgelegt ist und der Druckertreiber mehrere Kopien unterstützt, gibt das dmCopies-Element die Anzahl der vom Benutzer angeforderten Kopien an. Wenn dieses Kennzeichen festgelegt ist und der Druckertreiber die Sortierung unterstützt, gibt das dmCollate-Element der DEVMODE-Struktur an, ob der Benutzer die Sortierung wünscht. Wenn dieses Kennzeichen nicht festgelegt ist, gibt das dmCopies-Element immer 1 zurück, und das dmCollate-Element ist immer null.
'Bekanntes Problem unter Windows 2000/XP/2003: Wenn dieses Kennzeichen nicht festgelegt ist, bevor PrintDlg aufgerufen wird, tauschen PrintDlgmöglicherweise nCopies - und dmCopies-Werte , wenn sie zurückgegeben wird. Die Problemumgehung für dieses Problem wird dmCopies verwendet, wenn der Wert größer als 1 ist, andernfalls verwenden Sie nCopies, damit Sie die tatsächliche Anzahl der Kopien abrufen, die gedruckt werden sollen, wenn PrintDlg zurückgibt.
Private Const PD_DISABLEPRINTTOFILE         As Long = &H80000  '    Deaktiviert das Kontrollkästchen "In Datei drucken ".
Private Const PD_HIDEPRINTTOFILE            As Long = &H100000 ' cdlPDHidePrintToFile    = 1048576 (&H100000) ' Blendet das Kontrollkästchen "In Datei drucken " aus.
Private Const PD_NONETWORKBUTTON            As Long = &H200000 '    Blendet die Netzwerkschaltfläche aus und deaktiviert sie.

'Enum PrinterOrientationConstants
'    cdlPortrait = 1
'    cdlLandscape = 2



'Oh wowo if you deal with PrintDialog you come across a whole new universe of printer settings and print settings, paper size and stuff

'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646942(v=vs.85)
'https://docs.microsoft.com/de-de/windows/win32/api/commdlg/ns-commdlg-printdlgexw

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

'typedef struct tagPDA {
'  DWORD           lStructSize;         ' 4
'  HWND            hwndOwner;           ' 4 - 8
'  HGLOBAL         hDevMode;            ' 4 - 8
'  HGLOBAL         hDevNames;           ' 4 - 8
'  HDC             hDC;                 ' 4 - 8
'  DWORD           Flags;               ' 4
'  WORD            nFromPage;           ' 2
'  WORD            nToPage;             ' 2
'  WORD            nMinPage;            ' 2
'  WORD            nMaxPage;            ' 2
'  WORD            nCopies;             ' 2
'  HINSTANCE       hInstance;           ' 4 - 8
'  LPARAM          lCustData;           ' 4 - 8
'  LPPRINTHOOKPROC lpfnPrintHook;       ' 4 - 8
'  LPSETUPHOOKPROC lpfnSetupHook;       ' 4 - 8
'  LPCSTR          lpPrintTemplateName; ' 4 - 8
'  LPCSTR          lpSetupTemplateName; ' 4 - 8
'  HGLOBAL         hPrintTemplate;      ' 4 - 8
'  HGLOBAL         hSetupTemplate;      ' 4 - 8
'} PRINTDLGA, *LPPRINTDLGA;       ' Sum:

Private Type PrintDlg
    lStructSize   As Long
    hwndOwner     As LongPtr
    hdevmode      As LongPtr
    hDevNames     As LongPtr
    hDC           As LongPtr
    flags         As Long
    nFromPage     As Integer
    nToPage       As Integer
    nMinPage      As Integer
    nMaxPage      As Integer
    nCopies       As Integer
    hInstance     As LongPtr
    lCustData     As Long
    lpfnPrintHook As LongPtr
    lpfnSetupHook As LongPtr
    lpPrintTemplateName As LongPtr
    lpSetupTemplateName As LongPtr
    hPrintTemplate As LongPtr
    hSetupTemplate As LongPtr
End Type


'typedef struct tagPDEXW {
'  DWORD            lStructSize;
'  HWND             hwndOwner;
'  HGLOBAL          hDevMode;
'  HGLOBAL          hDevNames;
'  HDC              hDC;
'  DWORD            Flags;
'  DWORD            Flags2;
'  DWORD            ExclusionFlags;
'  DWORD            nPageRanges;
'  DWORD            nMaxPageRanges;
'  LPPRINTPAGERANGE lpPageRanges;
'  DWORD            nMinPage;
'  DWORD            nMaxPage;
'  DWORD            nCopies;
'  HINSTANCE        hInstance;
'  LPCWSTR          lpPrintTemplateName;
'  LPUNKNOWN        lpCallback;
'  DWORD            nPropertyPages;
'  HPROPSHEETPAGE   *lphPropertyPages;
'  DWORD            nStartPage;
'  DWORD            dwResultAction;
'} PRINTDLGEXW, *LPPRINTDLGEXW;

Private Type PrintDlgExW
    lStructSize      As Long
    hwndOwner        As LongPtr '  HWND
    hdevmode         As LongPtr '  HGLOBAL
    hDevNames        As LongPtr '  HGLOBAL
    hDC              As LongPtr '  HDC
    flags            As Long    '  DWORD
    Flags2           As Long    '  DWORD
    ExclusionFlags   As Long    '  DWORD
    nPageRanges      As Long    '  DWORD
    nMaxPageRanges   As Long    '  DWORD
    lpPageRanges     As LongPtr '  LPPRINTPAGERANGE
    nMinPage         As Long    '  DWORD
    nMaxPage         As Long    '  DWORD
    nCopies          As Long    '  DWORD
    hInstance        As LongPtr '  HINSTANCE
    lpPrintTemplName As LongPtr '  LPCWSTR
    lpCallback       As LongPtr '  LPUNKNOWN
    nPropertyPages   As Long    '  DWORD
    lphPropertyPages As LongPtr '  HPROPSHEETPAGE
    nStartPage       As Long    '  DWORD
    dwResultAction   As Long    '  DWORD
End Type

'typedef struct _PROPSHEETPAGEW_V4 {
'  LPCWSTR pszHeaderTitle;
'  LPCWSTR pszHeaderSubTitle;
'  HANDLE  hActCtx;
'  union {
'    HBITMAP hbmHeader;
'    LPCWSTR pszbmHeader;
'  } DUMMYUNIONNAME3;
'} PROPSHEETPAGEW_V4, *LPPROPSHEETPAGEW_V4;

Private Type PROPSHEETPAGEW_V4
    pszHeaderTitle    As LongPtr 'LPCWSTR
    pszHeaderSubTitle As LongPtr '  LPCWSTR
    hActCtx           As LongPtr '  HANDLE
'  union {
'    HBITMAP hbmHeader;
'    LPCWSTR pszbmHeader;
'  } DUMMYUNIONNAME3;
End Type

Private Declare Function PrintDlgA Lib "comdlg32" (ByRef pPrintdlg As PrintDlg) As Long

Private Declare Function PrintDlgA Lib "comdlg32" (ByRef pPrintdlg As PrintDlg) As Long

'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
Private Declare Function PrintDlgExW Lib "comdlg32" (ByRef TLPPRINTDLGEXW As PrintDlgExW) As Long
Private Declare Function PrinterMessageBoxW Lib "winspool.drv" (ByVal hPrinter As Long, ByVal error As Long, ByVal hwnd As Long, ByVal pText As String, ByVal pCaption As String, ByVal dwType As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private m_UseEXDialog As Boolean
Private m_PrintDlg    As PrintDlg
Private m_PrintDlgEx  As PrintDlgExW

Private Sub Class_Initialize()
    m_PrintDlg.lStructSize = LenB(m_PrintDlg)
    m_PrintDlgEx.lStructSize = LenB(m_PrintDlgEx)
End Sub

'Properties of Windows.Forms.PrintDialog
Public Property Get AllowCurrentPage() As Boolean
    'm_PrintDlg.flags and
End Property
Public Property Get AllowPrintToFile() As Boolean
    AllowPrintToFile = m_PrintDlg.flags And PD_DISABLEPRINTTOFILE = 0
End Property
Public Property Let AllowPrintToFile(ByVal Value As Boolean)
    AllowPrintToFile = m_PrintDlg.flags Or PD_DISABLEPRINTTOFILE
End Property

Public Property Get AllowSelection() As Boolean
    
End Property
Public Property Get AllowSomePages() As Boolean
    '
End Property

Public Property Get UseEXDialog() As Boolean
    UseEXDialog = m_UseEXDialog
End Property
Public Property Let UseEXDialog(ByVal Value As Boolean)
    m_UseEXDialog = Value
End Property

Public Property Get ShowNetwork() As Boolean
    ShowNetwork = (m_PrintDlg.flags And PD_NONETWORKBUTTON) = 0
End Property
'Public Property Get Container() As IContainer
'    '
'End Property
'Public Property Get Document_DefaultPageSettings() As PageSettings
'    '
'End Property
Public Property Get Document_DocumentName() As String
    '
End Property
Public Property Get Document_OriginAtMargins() As Boolean
    '
End Property
Public Sub Document_Print()
    '
End Sub
Public Property Get Document_PrintController_IsPreview() As Boolean
    '
End Property
'Public Property Get Document_PrinterSettings() As PrinterSettings
'    '
'End Property
'Public Property Get Document_Site() As ISite
'    '
'End Property
Public Function Document_ToString() As String
    '
End Function
'Public Property Get PrinterSettings() As PrinterSettings
'    '
'End Property
Public Property Get PrinterSettings_CanDuplex() As Boolean
    '
End Property
Public Property Get PrinterSettings_Collate() As Boolean
    '
End Property
Public Property Get PrinterSettings_Copies() As Integer
    '
End Property
'Public Property Get PrinterSettings_DefaultPageSettings_Bounds() As Rectangle
'    '
'End Property
Public Property Get PrinterSettings_DefaultPageSettings_Clone() As Object
    '
End Property
Public Property Get PrinterSettings_DefaultPageSettings_Color() As Boolean
    '
End Property
'Public Property Get PrinterSettings_DefaultPageSettings_CopyToHdevmode(hdevmode As IntPtr)
'    '
'End Property
'Public Property Get PrinterSettings_Duplex() As Duplex 'Enum
'    '
'End Property
Public Property Get PrinterSettings_Duplex_CompareTo()
    '
End Property
Public Property Get PrinterSettings_Duplex_Default() As Integer
    '
End Property
Public Property Get PrinterSettings_Duplex_HasFlag(flag As Long) As Boolean
    '
End Property
Public Property Get PrinterSettings_Duplex_Horizontal()
    'Horizontal = 3
End Property
Public Property Get PrinterSettings_Duplex_Simplex()
    'Simplex = 1
End Property
Public Property Get PrinterSettings_Duplex_Vertical()
    'Vertical = 2
End Property
Public Function PrinterSettings_Duplex_ToString()
    '
End Function


Public Property Get PrinterSettings_FromPage() As Integer
    '
End Property
Public Property Get PrinterSettings_IsDefaultPrinter() As Boolean
    '
End Property
Public Property Get PrinterSettings_IsDirectPrintingSupported() As Boolean
    '
End Property
Public Property Get PrinterSettings_IsPlotter() As Boolean
    '
End Property
Public Property Get PrinterSettings_IsValid() As Boolean
    '
End Property
Public Property Get PrinterSettings_LandscapeAngle() As Long
    '
End Property
Public Property Get PrinterSettings_MaximumCopies() As Long
    '
End Property
Public Property Get PrinterSettings_MaximumPage() As Long
    '
End Property
Public Property Get PrinterSettings_MinimumPage() As Long
    '
End Property
'Public Property Get PrinterSettings_PaperSizes() As PaperSizeCollection
'    '
'End Property
'Public Property Get PrinterSettings_PaperSources() As PaperSourceCollection
'    '
'End Property
Public Property Get PrinterSettings_PrinterName() As String
    '
End Property
'Public Property Get PrinterSettings_PrinterResolutions() As PrinterResolutionCollection
'    '
'End Property
Public Property Get PrinterSettings_PrintFileName() As String
    '
End Property
'Public Property Get PrinterSettings_PrintRange() As PrintRange
'    '
'End Property
Public Property Get PrinterSettings_PrintToFile() As Boolean
    '
End Property
Public Property Get PrinterSettings_SupportsColor() As Boolean
    '
End Property
Public Property Get PrinterSettings_ToPage() As Long
    '
End Property
Public Property Get PrintToFile() As Boolean
    '
End Property
Public Sub Reset()
    '
End Sub

Public Function ShowDialog() As VbMsgBoxResult

'https://microsoft.public.de.vb.narkive.com/wz5CVka1/druckereinstellung-mit-printdialog-oder-commondialog-showprinter
Try: On Error GoTo Catch
    Dim hr As Long
    If m_UseEXDialog Then
        hr = PrintDlgExW(m_PrintDlgEx)
    Else
        m_PrintDlg.hwndOwner = FMain.hwnd
        m_PrintDlg.flags = PD_USEDEVMODECOPIESANDCOLLATE Or PD_RETURNDC
        m_PrintDlg.nCopies = 1
        m_PrintDlg.nFromPage = 10 '&HFFFF
        m_PrintDlg.nToPage = 2 '&HFFFF
        m_PrintDlg.nMinPage = 1
        m_PrintDlg.nMaxPage = 10 '&HFFFF
        hr = PrintDlgA(m_PrintDlg)
    End If
    ShowDialog = IIf(hr, vbOK, vbCancel)
    Exit Function
Catch:
    MsgBox "error by calloing PrintDlg"
End Function

Public Property Get ShowHelp() As Boolean
    '
End Property
'Public Property Get ShowNetwork() As Boolean
'    '
'End Property
Public Property Get Site()
    '
End Property
Public Property Get Tag()
    '
End Property
Public Function ToString() As String
    '
End Function
'Public Property Get UseEXDialog() As Boolean
'    '
'End Property
'
