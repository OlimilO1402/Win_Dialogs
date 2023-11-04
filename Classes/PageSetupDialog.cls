VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageSetupDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type POINTAPI
    X As Long
    y As Long
End Type
 
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
 
Private Type TPageSetupDlg
    lStructSize As Long
    hwndOwner   As LongPtr
    hDevMode    As LongPtr
    hDevNames   As LongPtr
    flags       As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin    As RECT
    hInstance   As LongPtr
    lCustData   As Long
    lpfnPageSetupHook       As LongPtr
    lpfnPagePaintHook       As LongPtr
    lpPageSetupTemplateName As LongPtr 'String Pointer
    hPageSetupTemplate      As LongPtr
End Type
'
'PageSetupDialog erfordert, dass ein PageSettings-Objekt angezeigt wird.
'Legen Sie PageSetupDialog.Document (Bevorzugt),
'PageSetupDialog.PrinterSettings oder PageSetupDialog.PageSettings fest

Private Type TDeviceNames ' DEVNAMES
    wDriverOffset  As Integer
    wDeviceOffset  As Integer
    wOutputOffset  As Integer
    wDefault       As Integer
    extra(0 To 2 * 100 - 1) As Byte '* 100
End Type

Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32

Private Type TDeviceMode
    dmDeviceName(0 To CCHDEVICENAME * 2 - 1) As Byte 'String * CCHDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    
    PaperUnion(0 To 15) As Byte
    
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName(0 To CCHFORMNAME * 2 - 1) As Byte 'As String * CCHFORMNAME
    
    dmLogPixels        As Integer
    dmBitsPerPel       As Long
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    
    DisplayUnion(0 To 31) As Byte
    
End Type
Private Type TPaperUnion1
    dmOrientation      As Integer ' 2
    dmPaperSize        As Integer ' 2
    dmPaperLength      As Integer ' 2
    dmPaperWidth       As Integer ' 2
    dmScale            As Integer ' 2
    dmCopies           As Integer ' 2
    dmDefaultSource    As Integer ' 2
    dmPrintQuality     As Integer ' 2
End Type                    ' Sum: 16
Private Type TPaperUnion2
    dmPosition           As POINTAPI ' 8
    dmDisplayOrientation As Long     ' 4
    dmDisplayFixedOutput As Long     ' 4
End Type                       ' Sum: 16
Private Type TDisplayUnion1
    dmDisplayFlags     As Long '  4
    dmNup              As Long '  4
    dmReserved(0 To 5) As Long ' 24
End Type                  ' Sum: 32
Private Type TDisplayUnion2
    dmDisplayFrequency As Long ' 4
    dmICMMethod        As Long ' 4
    dmMediaType        As Long ' 4
    dmDitherType       As Long ' 4
    dmReserved1        As Long ' 4
    dmReserved2        As Long ' 4
    dmPanningWidth     As Long ' 4
    dmPanningHeight    As Long ' 4
End Type                 ' Sum: 32

'typedef struct _devicemodeW {
'  WCHAR dmDeviceName[CCHDEVICENAME];
'  WORD  dmSpecVersion;
'  WORD  dmDriverVersion;
'  WORD  dmSize;
'  WORD  dmDriverExtra;
'  DWORD dmFields;
'  union {
'    struct {
'      short dmOrientation;
'      short dmPaperSize;
'      short dmPaperLength;
'      short dmPaperWidth;
'      short dmScale;
'      short dmCopies;
'      short dmDefaultSource;
'      short dmPrintQuality;
'    } DUMMYSTRUCTNAME;
'    POINTL dmPosition;
'    struct {
'      POINTL dmPosition;
'      DWORD  dmDisplayOrientation;
'      DWORD  dmDisplayFixedOutput;
'    } DUMMYSTRUCTNAME2;
'  } DUMMYUNIONNAME;
'  short dmColor;
'  short dmDuplex;
'  short dmYResolution;
'  short dmTTOption;
'  short dmCollate;
'  WCHAR dmFormName[CCHFORMNAME];
'  WORD  dmLogPixels;
'  DWORD dmBitsPerPel;
'  DWORD dmPelsWidth;
'  DWORD dmPelsHeight;
'  union {
'    DWORD dmDisplayFlags;
'    DWORD dmNup;
'  } DUMMYUNIONNAME2;
'  DWORD dmDisplayFrequency;
'  DWORD dmICMMethod;
'  DWORD dmICMIntent;
'  DWORD dmMediaType;
'  DWORD dmDitherType;
'  DWORD dmReserved1;
'  DWORD dmReserved2;
'  DWORD dmPanningWidth;
'  DWORD dmPanningHeight;
'} DEVMODEW, *PDEVMODEW, *NPDEVMODEW, *LPDEVMODEW;

'PAGESETUPDLG flags-Konstanten
Private Const PSD_DEFAULTMINMARGINS             As Long = &H0        ' Legt die Mindestwerte fest, die der Benutzer für die Seitenränder angeben kann, um die vom Drucker zulässigen Mindestränder zu sein. Dies ist die Standardoption. Dieses Flag wird ignoriert, wenn auch die flags PSD_MARGINS und PSD_MINMARGINS angegeben sind.
Private Const PSD_INWININIINTLMEASURE           As Long = &H0        ' Reserviert.

Private Const PSD_MINMARGINS                    As Long = &H1        ' Bewirkt, dass das System die im rtMinMargin-Element angegebenen Werte als mindest zulässige Breite für den linken, oberen, rechten und unteren Rand verwendet. Das System verhindert, dass der Benutzer eine Breite eingibt, die kleiner als der angegebene Mindestwert ist. Wenn PSD_MINMARGINS nicht angegeben ist, legt das System die vom Drucker zulässigen Mindestbreiten fest.
Private Const PSD_MARGINS                       As Long = &H2        ' Bewirkt, dass das System die im rtMargin-Element angegebenen Werte als Anfangsbreiten für den linken, oberen, rechten und unteren Rand verwendet. Wenn PSD_MARGINS nicht festgelegt ist, legt das System die Anfangsbreite für alle Ränder auf einen Zoll fest.
Private Const PSD_INTHOUSANDTHSOFINCHES         As Long = &H4        ' Gibt an, dass Tausendstel Zoll die Maßeinheit für Ränder und Papierformat sind. Die Werte in den Membern rtMargin, rtMinMargin und ptPaperSize sind in tausendstel Zoll angegeben. Sie können dieses Flag bei der Eingabe festlegen, um die Standardmesseinheit für das Gebietsschema des Benutzers außer Kraft zu setzen. Wenn die Funktion zurückgegeben wird, legt das Dialogfeld dieses Flag fest, um die verwendeten Einheiten anzugeben.
Private Const PSD_INHUNDREDTHSOFMILLIMETERS     As Long = &H8        ' Gibt an, dass Hundertstel Millimeter die Maßeinheit für Ränder und Papierformat sind. Die Werte in den Elementen rtMargin, rtMinMargin und ptPaperSize sind in hundertstel Millimetern angegeben. Sie können dieses Flag bei der Eingabe festlegen, um die Standardmesseinheit für das Gebietsschema des Benutzers außer Kraft zu setzen. Wenn die Funktion zurückgegeben wird, legt das Dialogfeld dieses Flag fest, um die verwendeten Einheiten anzugeben.

Private Const PSD_DISABLEMARGINS                As Long = &H10       '  Deaktiviert die Randsteuerelemente, hindert den Benutzer daran, die Ränder festzulegen.
Private Const PSD_DISABLEPRINTER                As Long = &H20       ' Veraltet. Windows XP/2000: Deaktiviert die Schaltfläche Drucker und verhindert, dass der Benutzer ein Dialogfeld aufruft, das zusätzliche Informationen zur Druckereinrichtung enthält.

Private Const PSD_NOWARNING                     As Long = &H80       ' Verhindert, dass das System eine Warnmeldung anzeigt, wenn kein Standarddrucker vorhanden ist.


Private Const PSD_DISABLEORIENTATION            As Long = &H100      '  Deaktiviert die Ausrichtungssteuerelemente, hindert den Benutzer daran, die Seitenausrichtung festzulegen.
Private Const PSD_DISABLEPAPER                  As Long = &H200      ' Deaktiviert die Papiersteuerelemente, hindert den Benutzer daran, Seitenparameter wie Papierformat und Quelle festzulegen.
Private Const PSD_RETURNDEFAULT                 As Long = &H400      ' PageSetupDlg zeigt das Dialogfeld nicht an. Stattdessen werden die Elemente hDevNames und hDevMode für DEVMODE - und DEVNAMES-Strukturen festgelegt, die für den Standarddrucker des Systems initialisiert werden. PageSetupDlg gibt einen Fehler zurück, wenn hDevNames oder hDevMode nicht NULL ist.
Private Const PSD_SHOWHELP                      As Long = &H800      '

Private Const PSD_ENABLEPAGESETUPHOOK           As Long = &H2000     '  Aktiviert die hook-Prozedur, die im lpfnPageSetupHook-Element angegeben ist.
Private Const PSD_ENABLEPAGESETUPTEMPLATE       As Long = &H8000     ' Gibt an, dass die Elemente hInstance und lpPageSetupTemplateName eine Dialogfeldvorlage angeben, die anstelle der Standardvorlage verwendet werden soll.

Private Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE As Long = &H20000    ' Gibt an, dass das hPageSetupTemplate-Element einen Datenblock identifiziert, der eine vorab geladene Dialogfeldvorlage enthält. Das System ignoriert den lpPageSetupTemplateName-Member , wenn dieses Flag angegeben ist.
Private Const PSD_ENABLEPAGEPAINTHOOK           As Long = &H40000    '  Aktiviert die hook-Prozedur, die im lpfnPagePaintHook-Member angegeben ist.
Private Const PSD_DISABLEPAGEPAINTING           As Long = &H80000    ' Verhindert, dass das Dialogfeld den Inhalt der Beispielseite zeichnet. Wenn Sie eine PagePaintHook-Hookprozedur aktivieren, können Sie weiterhin den Inhalt der Beispielseite zeichnen.

Private Const PSD_NONETWORKBUTTON               As Long = &H200000   ' Blendet die Schaltfläche Netzwerk aus und deaktiviert sie.


#If VBA7 Then
    Private Declare PtrSafe Function PageSetupDlgW Lib "comdlg32" (pPagesetupdlg As TPageSetupDlg) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLength As Long)
#Else
    Private Declare Function PageSetupDlgW Lib "comdlg32" (pPagesetupdlg As TPageSetupDlg) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLength As Long)
#End If

Private m_DriverName  As String
Private m_DeviceName  As String
Private m_OutputName  As String
Private m_DefaultName As String

Private m_PSD  As TPageSetupDlg
Private mFlags As Long

Private Sub Class_Initialize()
    With m_PSD
        '.hInstance = App.hInstance
        '.hwndOwner
        .lStructSize = LenB(m_PSD)
        '.flags = .flags Or PSD_INHUNDREDTHSOFMILLIMETERS
        mFlags = mFlags Or PSD_INHUNDREDTHSOFMILLIMETERS
    End With
End Sub

Private Sub Class_Terminate()
    'Speicher befreien von den Strukturen
    With m_PSD
        GlobalFree .hDevMode
        GlobalFree .hDevNames
    End With
End Sub

Private Function GethWnd(Frm As Variant) As LongPtr
Try: On Error GoTo Catch
    GethWnd = Frm.hwnd
    Exit Function
Catch:
    GethWnd = GetActiveWindow
End Function

Private Function App_hInstance() As LongPtr
    'App_hInstance = GetModuleHandleW(StrPtr("WINWORD.exe"))
    'App_hInstance = Excel.Application.hInstance
    App_hInstance = GetModuleHandleW(0)
End Function

Private Function GethInst() As LongPtr
Try: On Error GoTo Finally
#If VBA6 Or VBA7 Then
    GethInst = App_hInstance 'Ptr
#Else
    GethInst = App.hInstance
#End If
Finally:
    GethInst = GetModuleHandleW(0) ' in VBA7 correct is also: (0^)
End Function

Public Function ShowDialog(Optional Frm As Variant) As MsgBoxResult

Try: On Error GoTo Catch
    
    Dim aOwnerhWnd  As LongPtr: aOwnerhWnd = GethWnd(Frm)
    Dim aOwnerhInst As LongPtr: aOwnerhInst = GethInst
    With m_PSD
        .hwndOwner = aOwnerhWnd
        .hInstance = aOwnerhInst
        .flags = mFlags
    End With
    Dim hr As Long: hr = PageSetupDlgW(m_PSD)
    If hr Then
        ShowDialog = vbOK
        GetNames m_PSD.hDevNames
    Else
        ShowDialog = vbCancel
    End If
    Exit Function
Catch:
    If Err.Number = 32755 Then
        ShowDialog = vbCancel
    Else
        ShowDialog = vbAbort
        MsgBox CStr(Err.Number) + " " + Err.Description
    End If
End Function

'Beispiel: Druckernamen der gewählten Einstellungen ermitteln
' (Geht so auch bei hDevMode->DevMode-Struktur)
Private Sub GetNames(ByVal hDevNames As LongPtr) 'As String
    
    'Struktur anhand des Handles füllen (Kopieren)
    Dim LenDevNames  As Long: LenDevNames = GlobalSize(hDevNames)
    Dim pDevNames As LongPtr:   pDevNames = GlobalLock(hDevNames)
    Dim DN As TDeviceNames 'DEVNAMES
    RtlMoveMemory DN, ByVal pDevNames, LenDevNames
    GlobalUnlock hDevNames
    
    'Struktur auswerten
    With DN
        Dim ex As String: ex = .extra
        Dim sa() As String: sa = Split(ex, vbNullChar)
        Dim i As Long
        If Len(sa(i)) Then m_DriverName = sa(i)
        i = i + 1
        If Len(sa(i)) Then m_DeviceName = sa(i)
        i = i + 1
        If Len(sa(i)) Then m_OutputName = sa(i)
        'm_DriverName =
        'm_DriverName = Left$(.extra, .wDriverOffset)
        'mDeviceName = Mid$(.extra, .wDriverOffset + 3, .wDeviceOffset - .wDriverOffset)
        'm_OutputName = Mid$(.extra, .wDeviceOffset + 3, .wOutputOffset - .wDeviceOffset)
        
        'vvv die formel ist von VB-API-Helpline
        'm_DeviceName = Mid$(.extra, .wDeviceOffset - .wDriverOffset + 1, .wOutputOffset - .wDeviceOffset - 2)
        'Debug.Print "wDeviceOffset: " & .wDeviceOffset
        'Debug.Print "wDriverOffset: " & .wDriverOffset
        'Debug.Print "wOutputOffset: " & .wOutputOffset
        'Debug.Print "wDefault:      " & .wDefault
        'Debug.Print .extra
    End With
End Sub

Public Property Get DriverName() As String
    DriverName = m_DriverName
End Property

Public Property Get DeviceName() As String
    DeviceName = m_DeviceName
End Property

Public Property Get OutputName() As String
    OutputName = m_OutputName
End Property

Public Property Get PaperSizeWidth() As Long
    PaperSizeWidth = m_PSD.ptPaperSize.X
End Property

Public Property Get PaperSizeHeight() As Long
    PaperSizeHeight = m_PSD.ptPaperSize.y
End Property

Public Property Get MarginsLeft() As Long
    MarginsLeft = m_PSD.rtMargin.Left
End Property

Public Property Get MarginsRight() As Long
    MarginsRight = m_PSD.rtMargin.Right
End Property

Public Property Get MarginsTop() As Long
    MarginsTop = m_PSD.rtMargin.Top
End Property

Public Property Get MarginsBottom() As Long
    MarginsBottom = m_PSD.rtMargin.Bottom
End Property

Public Property Get MinMarginsLeft() As Long
    MinMarginsLeft = m_PSD.rtMinMargin.Left
End Property

Public Property Get MinMarginsRight() As Long
    MinMarginsRight = m_PSD.rtMinMargin.Right
End Property

Public Property Get MinMarginsTop() As Long
    MinMarginsTop = m_PSD.rtMinMargin.Top
End Property

Public Property Get MinMarginsBottom() As Long
    MinMarginsBottom = m_PSD.rtMinMargin.Bottom
End Property

' v ############################## v '    Flags    ' v ############################## v '
Public Property Get AllowMargins() As Boolean
    AllowMargins = Not CBool(mFlags And PSD_DISABLEMARGINS)
End Property
Public Property Let AllowMargins(ByVal Value As Boolean)
    mFlags = mFlags Or PSD_MARGINS
    If Value Then mFlags = mFlags Xor PSD_DISABLEMARGINS
End Property

Public Property Get AllowOrientation() As Boolean
    AllowOrientation = Not CBool(mFlags And PSD_DISABLEORIENTATION)
End Property
Public Property Let AllowOrientation(ByVal Value As Boolean)
    mFlags = mFlags Or PSD_DISABLEORIENTATION
    If Value Then mFlags = mFlags Xor PSD_DISABLEORIENTATION
End Property

Public Property Get AllowPaper() As Boolean
    AllowPaper = Not CBool(mFlags And PSD_DISABLEPAPER)
End Property
Public Property Let AllowPaper(ByVal Value As Boolean)
    mFlags = mFlags Or PSD_DISABLEPAPER
    If Value Then mFlags = mFlags Xor PSD_DISABLEPAPER
End Property

Public Property Get AllowPrinter() As Boolean
    AllowPrinter = Not CBool(mFlags And PSD_DISABLEPRINTER)
End Property
Public Property Let AllowPrinter(ByVal Value As Boolean)
    mFlags = mFlags Or PSD_DISABLEPRINTER
    If Value Then mFlags = mFlags Xor PSD_DISABLEPRINTER
End Property

Public Property Get ShowHelp() As Boolean
    ShowHelp = mFlags And PSD_SHOWHELP
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    mFlags = mFlags Or PSD_SHOWHELP
    If Not Value Then mFlags = mFlags Xor PSD_SHOWHELP
End Property

Public Property Get EnableMetric() As Boolean
    'EnableMetric= mFlags and
End Property
Public Property Let EnableMetric(ByVal Value As Boolean)
    'mFlags = mFlags Or PSD_MARGINS
    'If Not Value Then mFlags = mFlags Xor PSD_MARGINS
End Property

Public Property Get ShowNetwork() As Boolean
    'ShowNetwork= mFlags and
End Property
Public Property Let ShowNetwork(ByVal Value As Boolean)
    'mFlags = mFlags Or PSD_MARGINS
    'If Not Value Then mFlags = mFlags Xor PSD_MARGINS
End Property

'psd.MinMargins.Left
'psd.MinMargins.Right
'psd.MinMargins.Top
'psd.MinMargins.Bottom
'psd.MinMargins.

'psd.Document As System.Drawing.Printing.PrintDocument
'psd.PageSettings As System.Drawing.Printing.PageSettings
'psd.PrinterSettings As System.Drawing.Printing.PrinterSettings

