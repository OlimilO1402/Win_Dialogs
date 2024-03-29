VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CF_SCREENFONTS          As Long = &H1&      ' Listet Bildschirm-Fonts
Private Const CF_PRINTERFONTS         As Long = &H2&      ' Listet Drucker-Fonts
Private Const CF_BOTH                 As Long = &H3&      ' Listet Drucker- und Bildschirm-Fonts
Private Const CF_SHOWHELP             As Long = &H4&      ' Zeigt den Hilfe-Button an
Private Const CF_ENABLEHOOK           As Long = &H8&      ' Aktiviert die Callback-Funktion
Private Const CF_ENABLETEMPLATE       As Long = &H10&     ' Der Dialog benutzt Template's die von' TemplateNames festgelegt sind
Private Const CF_ENABLETEMPLATEHANDLE As Long = &H20&     ' Verwendet den durch hInstance festgelegten Dialog
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&     ' Setzt die Startwerte, welche �ber die' LOGFONT-Struktur angegeben wurden
Private Const CF_USESTYLE             As Long = &H80&     ' Verwendet die in "lpStyle" angegebenen Werte
Private Const CF_EFFECTS              As Long = &H100&    ' Erlaubt Font-Eigenschaften wie' Unterstreichen, Farbe und Durchgestrichen
Private Const CF_APPLY                As Long = &H200&    ' Zeigt einen "�bernehmen" Button an
Private Const CF_ANSIONLY             As Long = &H400&    ' Zeigt nur Windows- oder Unicode-Fonts
Private Const CF_SCRIPTSONLY          As Long = &H400&    ' Listet nur Windows- oder Unicode-Fonts
Private Const CF_NOOEMFONTS           As Long = &H800&    ' Zeigt keine OEM Fonts
Private Const CF_NOVECTORFONTS        As Long = &H800&    ' keine Vector-Fonts anzeigen
Private Const CF_NOSIMULATIONS        As Long = &H1000&   ' Kein Beispiel (Vorschau) anzeigen
Private Const CF_LIMITSIZE            As Long = &H2000&   ' Erlaubt nur Schriftgr��en im Bereich "nSizeMin" und "nSizeMax"
Private Const CF_FIXEDPITCHONLY       As Long = &H4000&   ' Listet nur Fixed-Pitch Fonts
Private Const CF_WYSIWYG              As Long = &H8000&   ' Listet nur Fonts, die Drucker- und Bildschirm-Fonts gleichzeitig sind' (muss benutzt werden mit CF_BOTH und CF_SCALABLEONLY)
Private Const CF_FORCEFONTEXIST       As Long = &H10000   ' Verweigert die Eingabe nicht aufgef�hrter Fonts
Private Const CF_SCALABLEONLY         As Long = &H20000   ' Listet nur skalierbare Fonts
Private Const CF_TTONLY               As Long = &H40000   ' Listet nur TrueType-Schriftarten
Private Const CF_NOFACESEL            As Long = &H80000   ' Kein Standard Facenamen selektieren
Private Const CF_NOSTYLESEL           As Long = &H100000  ' kein Standard-Stil setzen
Private Const CF_NOSIZESEL            As Long = &H200000  ' keine Standardgr��e setzen
Private Const CF_SELECTSCRIPT         As Long = &H400000  ' Listet nur Script-Fonts
Private Const CF_NOSCRIPTSEL          As Long = &H800000  ' Kein Standard Script selektieren
Private Const CF_NOVERTFONTS          As Long = &H1000000 ' keine vertikal ausgerichteten Fonts anzeigen

' nFontType Konstanten
Private Const BOLD_FONTTYPE       As Long = &H100&   ' fett
Private Const ITALIC_FONTTYPE     As Long = &H200&   ' kursiv
Private Const REGULAR_FONTTYPE    As Long = &H400&   ' regul�re Fonts
Private Const SCREEN_FONTTYPE     As Long = &H2000&  ' Bildschirm-Fonts
Private Const PRINTER_FONTTYPE    As Long = &H4000&  ' Drucker-Fonts
Private Const SIMULATED_FONTTYPE  As Long = &H8000&  ' Fonts, die in der Dialogbox simuliert werden k�nnen

' lfWeight Konstanten
Private Const FW_DONTCARE         As Long = 0&    ' Standard
Private Const FW_THIN             As Long = 100&  ' super d�nn
Private Const FW_EXTRALIGHT       As Long = 200&  ' extra d�nn
Private Const FW_LIGHT            As Long = 300&  ' d�nn
Private Const FW_NORMAL           As Long = 400&  ' normal
Private Const FW_MEDIUM           As Long = 500&  ' mittel
Private Const FW_SEMIBOLD         As Long = 600&  ' etwas dicker
Private Const FW_BOLD             As Long = 700&  ' fett
Private Const FW_EXTRABOLD        As Long = 800&  ' extra fett
Private Const FW_HEAVY            As Long = 900&  ' super fett
Private Const FW_BLACK            As Long = 900&  ' super fett

' lfCharSet Konstanten 'Skript
Private Const ANSI_CHARSET        As Long = 0     ' Ansi Zeichensatz
Private Const DEFAULT_CHARSET     As Long = 1     ' Standard
Private Const SYMBOL_CHARSET      As Long = 2     ' Symbolisch
Private Const MAC_CHARSET         As Long = 77    ' Mac (Win 9x)
Private Const SHIFTJIS_CHARSET    As Long = 128   ' ShiftJis
Private Const HANGEUL_CHARSET     As Long = 129   ' Handgeul
Private Const JOHAB_CHARSET       As Long = 130   ' Johab (Win 9x)
Private Const GB2312_CHARSET      As Long = 134   ' Englisch
Private Const CHINESEBIG5_CHARSET As Long = 136   ' Chinesisch
Private Const GREEK_CHARSET       As Long = 161   ' Griechisch (Win 9x)
Private Const TURKISH_CHARSET     As Long = 162   ' T�rkisch (Win 9x)
Private Const HEBREW_CHARSET      As Long = 177   ' Hebr�isch (NT/2000)
Private Const ARABIC_CHARSET      As Long = 178   ' Arabisch (NT/2000)
Private Const BALTIC_CHARSET      As Long = 186   ' Baltisch (Win 9x)
Private Const RUSSIAN_CHARSET     As Long = 204   ' Russisch (Win 9x)
Private Const EASTEUROPE_CHARSET  As Long = 238   ' Osteurop�isch (Win 9x)
Private Const THAI_CHARSET        As Long = 222   ' Thail�ndisch (NT/2000)
Private Const OEM_CHARSET         As Long = 255   ' OEM

' lfOutPrecision Konstanten
Private Const OUT_DEFAULT_PRECIS  As Long = 0 ' Verwendet Standard-Fonts
Private Const OUT_STRING_PRECIS   As Long = 1 ' Raster-Font f�r die Enumeration
Private Const OUT_STROKE_PRECIS   As Long = 3 ' (Win 9x) Vector-Font f�r die Enumeration' (NT/2000) TrueType, Outline-Based oder VectorFont f�r die Enumeration
Private Const OUT_TT_PRECIS       As Long = 4 ' Verwendet eine TrueType-Font, wenn mehrere' Dateien mit dem selben Namen existieren
Private Const OUT_DEVICE_PRECIS   As Long = 5 ' Verwendet eine Device-Font wenn mehrere Dateien mit dem selben Namen existieren
Private Const OUT_RASTER_PRECIS   As Long = 6 ' Verwendet ein Raster-Font, wenn mehrere Dateien mit dem selben Namen existieren
Private Const OUT_TT_ONLY_PRECIS  As Long = 7 ' Verwendet eine TrueType-Schriftart
Private Const OUT_OUTLINE_PRECIS  As Long = 8 ' (NT/2000) Verwendet eine TureType Schriftart oder Outline-Based Fonts.

' lfClipPrecision Konstanten
Private Const CLIP_DEFAULT_PRECIS As Long = 0   ' Standard
Private Const CLIP_STROKE_PRECIS  As Long = 2   ' Raster-Vector oder TrueType-Fonts f�r die Enumeration
Private Const CLIP_LH_ANGLES      As Long = 16  ' die Richtung von irgendwelchen Rotationen wird vom Koordinatensystem festgelegt (gegen den Uhrzeigersinn)
Private Const CLIP_EMBEDDED       As Long = 128 ' wird benutzt f�r eingebettete schreibgesch�tzte Fonts

' lfQuality Konstanten
Private Const DEFAULT_QUALITY        As Long = 0 ' Standard Qualit�t
Private Const DRAFT_QUALITY          As Long = 1 ' entspricht "Schnelldruck" bei Druckausgaben
Private Const PROOF_QUALITY          As Long = 2 ' entspricht "Sch�ndruck" bei Druckausgaben
Private Const NONANTIALIASED_QUALITY As Long = 3 ' (Win 9x, NT 4.0, 2000) Schriftart wird nicht mit Anitialisierung gezeichnet
Private Const ANTIALIASED_QUALITY    As Long = 4 ' (Win 9x, NT 4.0, 2000) Schriftart wird wenn m�glich immer mit Anitialisierung gezeichnet

' lfPitchAndFamily Konstanten
Private Const DEFAULT_PITCH  As Long = 0  ' Standard
Private Const FIXED_PITCH    As Long = 1  ' Fest
Private Const VARIABLE_PITCH As Long = 2  ' Variabel

Private Const FF_DONTCARE    As Long = 0  ' Egal
Private Const FF_ROMAN       As Long = 16 ' Roman
Private Const FF_SWISS       As Long = 32 ' Swiss
Private Const FF_MODERN      As Long = 48 ' Modern
Private Const FF_SCRIPT      As Long = 64 ' Script
Private Const FF_DECORATIVE  As Long = 80 ' Dekoriert

'possible errors in dialog FontDialog
Private Const CDERR_STRUCTSIZE      As Long = &H1
Private Const CDERR_INITIALIZATION  As Long = &H2
Private Const CDERR_NOTEMPLATE      As Long = &H3
Private Const CDERR_NOHINSTANCE     As Long = &H4
Private Const CDERR_LOADSTRFAILURE  As Long = &H5
Private Const CDERR_FINDRESFAILURE  As Long = &H6
Private Const CDERR_LOADRESFAILURE  As Long = &H7
Private Const CDERR_LOCKRESFAILURE  As Long = &H8
Private Const CDERR_MEMALLOCFAILURE As Long = &H9
Private Const CDERR_MEMLOCKFAILURE  As Long = &HA
Private Const CDERR_NOHOOK          As Long = &HB
Private Const CFERR_MAXLESSTHANMIN  As Long = &H2002&
Private Const CFERR_NOFONTS         As Long = &H2001&
Private Const CDERR_DIALOGFAILURE   As Long = &HFFFF&

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

'typedef struct tagLOGFONTW {
'  LONG  lfHeight;
'  LONG  lfWidth;
'  LONG  lfEscapement;
'  LONG  lfOrientation;
'  LONG  lfWeight;
'  BYTE  lfItalic;
'  BYTE  lfUnderline;
'  BYTE  lfStrikeOut;
'  BYTE  lfCharSet;
'  BYTE  lfOutPrecision;
'  BYTE  lfClipPrecision;
'  BYTE  lfQuality;
'  BYTE  lfPitchAndFamily;
'  WCHAR lfFaceName[LF_FACESIZE];
'} LOGFONTW, *PLOGFONTW, *NPLOGFONTW, *LPLOGFONTW;

Private Const LF_FACESIZE    As Long = 32
Private Type LOGFONTW
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Integer 'String * LF_FACESIZE '(1 To LF_FACESIZE) As Byte
End Type

'typedef struct tagCHOOSEFONTW {
'  DWORD        lStructSize;
'  HWND         hwndOwner;
'  HDC          hDC;
'  LPLOGFONTW   lpLogFont;
'  INT          iPointSize;
'  DWORD        Flags;
'  COLORREF     rgbColors;
'  LPARAM       lCustData;
'  LPCFHOOKPROC lpfnHook;
'  LPCWSTR      lpTemplateName;
'  HINSTANCE    hInstance;
'  LPWSTR       lpszStyle;
'  WORD         nFontType;
'  WORD         ___MISSING_ALIGNMENT__;
'  INT          nSizeMin;
'  INT          nSizeMax;
'} CHOOSEFONTW;

Private Type ChooseFontW
    lStructSize    As Long    '4
    hwndOwner      As LongPtr '8  caller's window handle
    hDC            As LongPtr '8  printer DC/IC or NULL
    lpLogFont      As LongPtr '8  LOGFONT  ' ptr. to a LOGFONT struct
    iPointSize     As Long    '4  10 size points of selected font
    flags          As Long    '4  enum. type flags
    rgbColors      As Long    '4  returned text color
    lCustData      As LongPtr '8  data passed to hook fn.
    lpfnHook       As LongPtr '8  ptr. to hook function
    lpTemplateName As LongPtr '8  String ' custom template name
    hInstance      As LongPtr '8  Instance handle of.EXE that
                              '   contains cust. dlg. template
    lpszStyle      As LongPtr '8  String ' return the style field here
                              '   must be LF_FACESIZE or bigger
    nFontType      As Integer '2  same value reported to the EnumFonts
                              '   call back with the extra FONTTYPE_
                              '   bits added
    MISSING_ALIGN  As Integer '2
    nSizeMin       As Long    '4  minimum pt size allowed &
    nSizeMax       As Long    '4  max pt size allowed if
                              '   CF_LIMITSIZE is used
    'PadAtTheEnd    As Long    '4
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
    Private Declare PtrSafe Function ChooseFontW Lib "comdlg32" (ByRef pChoosefont As ChooseFontW) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long
    Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As LongPtr
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByRef pv As Any)
#Else
    Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
    Private Declare Function ChooseFontW Lib "comdlg32" (ByRef pChoosefont As ChooseFontW) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare Function CoTaskMemAlloc Lib "Ole32" (ByVal cb As Long) As LongPtr
    Private Declare Sub CoTaskMemFree Lib "Ole32" (ByRef pv As Any)
#End If
Private Const LOGPIXELSY As Long = 90
Private mDPI      As Long
Private mPT2Px    As Long
Private mFlags    As Long
Private mColor    As Long
Private mFont     As LOGFONTW
Private mMinSize  As Long
Private mMaxSize  As Long
Private mTag      As Variant
'Private mFontName As String

'TODO OM: ChooseFont hier als Variable

Private Sub Class_Initialize()
    mDPI = GetDeviceCaps(GetDC(0), LOGPIXELSY)
    mPT2Px = 72
    Reset
End Sub

Private Function Trim0(ByVal s As String) As String
    Trim0 = VBA.Strings.Trim$(Left$(s, lstrlenW(ByVal StrPtr(s))))
End Function

'Public Overrides Sub Reset()
Public Sub Reset()
    Dim StdName As String: StdName = "Ms Sans Serif" '"Arial", "Verdana", "Tahoma", "Consolas" what'd you prefer?
    Me.FontName = StdName
    ' Setzt alle Optionen im Dialogfeld auf die Standardwerte zur�ck.
    mFlags = mFlags Or CF_BOTH
    mFlags = mFlags Or CF_EFFECTS
    mFlags = mFlags Or CF_INITTOLOGFONTSTRUCT
    mFont.lfHeight = S2H(8) '10 + 2 / 3
    mMaxSize = 32767
End Sub
Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Private Function S2H(aSize As Long) As Long
    S2H = -MulDiv(aSize, mDPI, mPT2Px)
End Function
Private Function H2S(aHeight As Long) As Long
    H2S = -MulDiv(aHeight, mPT2Px, mDPI)
End Function
'Public Property AllowScriptChange() As Boolean
Public Property Get AllowScriptChange() As Boolean
    ' Ruft einen Wert ab, der angibt, ob der Benutzer den im Kombinationsfeld Script
    ' angegebenen Zeichensatz �ndern kann, so dass anstelle des aktuellen ein anderer
    ' Zeichensatz angezeigt wird, oder legt diesen fest.
    AllowScriptChange = Not (mFlags And CF_NOSCRIPTSEL)
End Property
Public Property Let AllowScriptChange(ByVal Value As Boolean)
    mFlags = mFlags Or CF_NOSCRIPTSEL
    If Value Then mFlags = mFlags Xor CF_NOSCRIPTSEL
End Property

'Public Property AllowSimulations() As Boolean
Public Property Get AllowSimulations() As Boolean
    ' Ruft einen Wert ab, der angibt, ob GDI-Schriftartsimulationen
    ' (Graphics Device Interface) f�r das Dialogfeld zugelassen sind, oder legt diesen fest.
    'mit einfacheren Worten: Fontvorschau mit kurzem Schriftbeispiel
    AllowSimulations = Not (mFlags And CF_NOSIMULATIONS)
End Property
Public Property Let AllowSimulations(ByVal Value As Boolean)
    mFlags = mFlags Or CF_NOSIMULATIONS
    If Value Then mFlags = mFlags Xor CF_NOSIMULATIONS
End Property

'Public Property AllowVectorFonts() As Boolean
Public Property Get AllowVectorFonts() As Boolean
    ' Ruft einen Wert ab, der angibt, ob Vektorschriftarten im Dialogfeld ausgew�hlt werden
    ' k�nnen, oder legt diesen fest.
    AllowVectorFonts = Not (mFlags And CF_NOVECTORFONTS)
End Property
Public Property Let AllowVectorFonts(ByVal Value As Boolean)
    mFlags = mFlags Or CF_NOVECTORFONTS
    If Value Then mFlags = mFlags Xor CF_NOVECTORFONTS
End Property

'Public Property AllowVerticalFonts() As Boolean
Public Property Get AllowVerticalFonts() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld vertikale und horizontale Schriftarten
    ' oder nur horizontale Schriftarten angezeigt werden, oder legt diesen fest.
    AllowVerticalFonts = Not (mFlags And CF_NOVERTFONTS)
End Property
Public Property Let AllowVerticalFonts(ByVal Value As Boolean)
    mFlags = mFlags Or CF_NOVERTFONTS
    If Value Then mFlags = mFlags Xor CF_NOVERTFONTS
End Property

'Public Property Color() As System.Drawing.Color
Public Property Get Color() As Long
    ' Ruft die ausgew�hlte Schriftfarbe ab oder legt diese fest.
    Color = mColor
End Property
Public Property Let Color(LngVal As Long)
    mColor = LngVal
End Property

'Public Property FixedPitchOnly() As Boolean
Public Property Get FixedPitchOnly() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld nur Schriftarten mit fester
    ' Schriftbreite ausgew�hlt werden k�nnen, oder legt diesen fest.
    FixedPitchOnly = (mFlags And CF_FIXEDPITCHONLY)
End Property
Public Property Let FixedPitchOnly(ByVal Value As Boolean)
    mFlags = mFlags Or CF_FIXEDPITCHONLY
    If Not Value Then mFlags = mFlags Xor CF_FIXEDPITCHONLY
End Property

Public Property Get FontName() As String
    FontName = String(LF_FACESIZE, vbNullChar)
    RtlMoveMemory ByVal StrPtr(FontName), ByVal VarPtr(mFont.lfFaceName(0)), LF_FACESIZE * 2
    FontName = Trim(FontName)
End Property
Public Property Let FontName(ByVal Value As String)
    Dim lb As Long: lb = 2 * Min(LF_FACESIZE, Len(Value))
    RtlZeroMemory ByVal VarPtr(mFont.lfFaceName(0)), LF_FACESIZE * 2
    RtlMoveMemory ByVal VarPtr(mFont.lfFaceName(0)), ByVal StrPtr(Value), lb
End Property

#If VBA7 Then
Public Property Get Font() As NewFont
    Set Font = New NewFont
#Else
Public Property Get Font() As StdFont
    Set Font = New StdFont
#End If
    With Font
        .Bold = mFont.lfWidth
        .Charset = mFont.lfCharSet
        .Italic = mFont.lfItalic
        '.Name = StrConv(mFont.lfFaceName, vbUnicode)
        .Name = Me.FontName
        .Size = H2S(mFont.lfHeight)
        .Strikethrough = mFont.lfStrikeOut
        .Underline = mFont.lfUnderline
        .Weight = mFont.lfWeight
    End With
End Property

#If VBA7 Then
Public Property Set Font(aFont As NewFont)
#Else
Public Property Set Font(aFont As Font)
#End If
    With mFont
        .lfWidth = aFont.Bold
        .lfCharSet = aFont.Charset
        .lfItalic = aFont.Italic
        'Dim StdName As String: StdName = "Ms Sans Serif"
        '.lfFaceName = StrConv(FntVal.Name, vbFromUnicode) & vbNullChar
        
        .lfHeight = S2H(aFont.Size)
        .lfStrikeOut = aFont.Strikethrough
        .lfUnderline = aFont.Underline
        .lfWeight = aFont.Weight
    End With
    Me.FontName = aFont.Name
End Property

'Public Property FontMustExist() As Boolean
Public Property Get FontMustExist() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld eine Fehlerbedingung angegeben wird,
    ' wenn der Benutzer eine nicht vorhandene Schriftart oder ein nicht vorhandenes Format
    ' ausw�hlt, oder legt diesen fest.
    FontMustExist = (mFlags And CF_FORCEFONTEXIST)
End Property
Public Property Let FontMustExist(ByVal Value As Boolean)
    mFlags = mFlags Or CF_FORCEFONTEXIST
    If Not Value Then mFlags = mFlags Xor CF_FORCEFONTEXIST
End Property

'Public Property MaxSize() As Integer
Public Property Get MaxSize() As Long
    ' Ruft den gr��ten Schriftgrad ab, den der Benutzer ausw�hlen kann, oder legt diesen fest.
    MaxSize = mMaxSize
End Property
Public Property Let MaxSize(ByVal Value As Long)
    mMaxSize = Value
End Property

'Public Property MinSize() As Integer
Public Property Get MinSize() As Long
    ' Ruft den kleinsten Schriftgrad ab, den der Benutzer ausw�hlen kann, oder legt diesen fest.
    MinSize = mMinSize
End Property
Public Property Let MinSize(ByVal Value As Long)
    mMinSize = Value
End Property

'Public Property ScriptsOnly() As Boolean
Public Property Get ScriptsOnly() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld Schriftarten f�r alle
    ' Nicht-OEM-Zeichens�tze und Symbolzeichens�tze sowie der ANSI-Zeichensatz
    ' ausgew�hlt werden k�nnen, oder legt diesen fest.
    ScriptsOnly = (mFlags And CF_SELECTSCRIPT)
End Property
Public Property Let ScriptsOnly(ByVal Value As Boolean)
    mFlags = mFlags Or CF_SELECTSCRIPT
    If Not Value Then mFlags = mFlags Xor CF_SELECTSCRIPT
End Property

'Public Property ShowApply() As Boolean
Public Property Get ShowApply() As Boolean
    ' Ruft einen Wert ab, der angibt, ob das Dialogfeld die Schaltfl�che Apply enth�lt,
    ' oder legt diesen fest.
    ShowApply = (mFlags And CF_APPLY)
End Property
Public Property Let ShowApply(ByVal Value As Boolean)
    mFlags = mFlags Or CF_APPLY
    If Not Value Then mFlags = mFlags Xor CF_APPLY
End Property

'Public Property ShowColor() As Boolean
Public Property Get ShowColor() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld die Farbauswahl angezeigt wird, oder legt diesen fest.
    '
End Property
Public Property Let ShowColor(BolVal As Boolean)
    '
End Property

'Public Property ShowEffects() As Boolean
Public Property Get ShowEffects() As Boolean
    ' Ruft einen Wert ab, der angibt, ob das Dialogfeld Steuerelemente enth�lt, mit denen der
    ' Benutzer Optionen zum Durchstreichen, Unterstreichen und f�r die Textfarbe angeben kann
    ' oder legt diesen fest.
    ShowEffects = (mFlags And CF_EFFECTS)
End Property
Public Property Let ShowEffects(ByVal Value As Boolean)
    mFlags = mFlags Or CF_EFFECTS
    If Not Value Then mFlags = mFlags Xor CF_EFFECTS
End Property

'Public Property ShowHelp() As Boolean
Public Property Get ShowHelp() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld die Hilfeschaltfl�che angezeigt wird, oder legt diesen fest.
    ShowHelp = (mFlags And CF_SHOWHELP)
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    mFlags = mFlags Or CF_SHOWHELP
    If Not Value Then mFlags = mFlags Xor CF_SHOWHELP
End Property

'Protected ReadOnly Property Options() As Integer
Public Property Get Options() As Long
    Options = mFlags
End Property


'wie schon hinl�nglich bekannt, f�gt einfach irgendwas hinzu
'wird intern nicht verwendet, kann extern verwendet werden
'Public Property Get Tag() As Object
Public Property Get Tag() As Variant
    Tag = mTag
End Property
Public Property Let Tag(Value)
    mTag = Value
End Property
Public Property Set Tag(Value)
    Set mTag = Value
End Property

Public Function ToString() As String
    ToString = "MBO.Forms.FontDialog"
End Function

Private Function GethWnd(Frm As Variant) As LongPtr
Try: On Error GoTo Catch
    GethWnd = Frm.hwnd
    Exit Function
Catch:
    GethWnd = GetActiveWindow
End Function

Private Function GethInst() As LongPtr
Try: On Error GoTo Finally
#If VBA6 Or VBA7 Then
    GethInst = Application.hInstance 'Ptr
#Else
    GethInst = App.hInstance
#End If
Finally:
    GethInst = GetModuleHandleW(0) ' in VBA7 correct is also: (0^)
End Function

Private Function CoTask_GetLogFontPtr() As LongPtr
    Dim bytLen As Long: bytLen = LenB(mFont)
    Dim p As LongPtr: p = CoTaskMemAlloc(bytLen)
    RtlMoveMemory ByVal p, mFont, bytLen
    CoTask_GetLogFontPtr = p
End Function
Private Property Let pFont(ByVal p As LongPtr)
    Dim bytLen As Long: bytLen = LenB(mFont)
    RtlMoveMemory mFont, ByVal p, bytLen
End Property

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult
Try: On Error GoTo Finally
    Dim aOwnerhWnd  As LongPtr: aOwnerhWnd = GethWnd(Frm)
    Dim aOwnerhInst As LongPtr: aOwnerhInst = GethInst
    Dim CF As ChooseFontW
    With CF
        .lStructSize = LenB(CF)
        'Debug.Print .lStructSize '104
        .hwndOwner = aOwnerhWnd
        .hInstance = aOwnerhInst
        '.hDC = Printer.hDC
        .flags = mFlags
        .iPointSize = 10
        .lpLogFont = CoTask_GetLogFontPtr 'VarPtr(mFont)
        .rgbColors = mColor
        .nSizeMax = mMaxSize
        .nSizeMin = mMinSize
    End With
    ShowDialog = ChooseFontW(CF)
    If ShowDialog Then
        pFont = CF.lpLogFont
        mColor = CF.rgbColors
        mMinSize = CF.nSizeMin
        mMaxSize = CF.nSizeMax
    Else
        ShowDialog = vbAbort
        CheckMessError
    End If
Finally:
    CoTaskMemFree ByVal CF.lpLogFont
    If Err Then ShowDialog = vbAbort
End Function

Private Sub CheckMessError()
    Dim hr As Long: hr = CommDlgExtendedError
    If hr = 0 Then Exit Sub
    Dim s As String
    Select Case hr
    Case CDERR_STRUCTSIZE:      s = "STRUCTSIZE"
    Case CDERR_INITIALIZATION:  s = "INITIALIZATION"
    Case CDERR_NOTEMPLATE:      s = "NOTEMPLATE"
    Case CDERR_NOHINSTANCE:     s = "NOHINSTANCE"
    Case CDERR_LOADSTRFAILURE:  s = "LOADSTRFAILURE"
    Case CDERR_FINDRESFAILURE:  s = "FINDRESFAILURE"
    Case CDERR_LOADRESFAILURE:  s = "LOADRESFAILURE"
    Case CDERR_LOCKRESFAILURE:  s = "LOCKRESFAILURE"
    Case CDERR_MEMALLOCFAILURE: s = "MEMALLOCFAILURE"
    Case CDERR_MEMLOCKFAILURE:  s = "MEMLOCKFAILURE"
    Case CDERR_NOHOOK:          s = "NOHOOK"
    Case CDERR_DIALOGFAILURE:   s = "DIALOGFAILURE"
    
    'Errors special for FontDialog:
    Case CFERR_MAXLESSTHANMIN:  s = "MAXLESSTHANMIN"
    Case CFERR_NOFONTS:         s = "NOFONTS"
    
    End Select
    If Len(s) Then
        MsgBox s
    End If
End Sub




