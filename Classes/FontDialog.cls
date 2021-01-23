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
Private Declare Function ChooseFontA Lib "comdlg32.dll" (ByRef pChoosefont As ChooseFont) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function MulDiv Lib "Kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Const LF_FACESIZE As Long = 32
Private Type LOGFONT
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
  lfFaceName As String * LF_FACESIZE '(1 To LF_FACESIZE) As Byte
End Type
Private Type ChooseFont
  lStructSize    As Long
  hwndOwner      As Long ' caller's window handle
  hdc            As Long ' printer DC/IC or NULL
  lpLogFont      As Long 'LOGFONT  ' ptr. to a LOGFONT struct
  iPointSize     As Long ' 10 size points of selected font
  flags          As Long ' enum. type flags
  rgbColors      As Long ' returned text color
  lCustData      As Long ' data passed to hook fn.
  lpfnHook       As Long ' ptr. to hook function
  lpTemplateName As String ' custom template name
  hInstance      As Long ' Instance handle of.EXE that
  ' contains cust. dlg. template
  lpszStyle      As String ' return the style field here
  ' must be LF_FACESIZE or bigger
  nFontType      As Integer ' same value reported to the EnumFonts
  ' call back with the extra FONTTYPE_
  ' bits added
  MISSING_ALIGNMENT As Integer
  nSizeMin    As Long ' minimum pt size allowed &
  nSizeMax    As Long ' max pt size allowed if
  ' CF_LIMITSIZE is used
End Type
Private Const CF_SCREENFONTS          As Long = &H1 ' Listet Bildschirm-Fonts
Private Const CF_PRINTERFONTS         As Long = &H2 ' Listet Drucker-Fonts
Private Const CF_BOTH                 As Long = &H3 ' Listet Drucker- und Bildschirm-Fonts
Private Const CF_SHOWHELP             As Long = &H4 ' Zeigt den Hilfe-Button an
Private Const CF_ENABLEHOOK           As Long = &H8 ' Aktiviert die Callback-Funktion
Private Const CF_ENABLETEMPLATE       As Long = &H10 ' Der Dialog benutzt Template's die von' TemplateNames festgelegt sind
Private Const CF_ENABLETEMPLATEHANDLE As Long = &H20 ' Verwendet den durch hInstance festgelegten Dialog
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40 ' Setzt die Startwerte, welche über die' LOGFONT-Struktur angegeben wurden
Private Const CF_USESTYLE             As Long = &H80 ' Verwendet die in "lpStyle" angegebenen Werte
Private Const CF_EFFECTS              As Long = &H100 ' Erlaubt Font-Eigenschaften wie' Unterstreichen, Farbe und Durchgestrichen
Private Const CF_APPLY                As Long = &H200 ' Zeigt einen "Übernehmen" Button an
Private Const CF_ANSIONLY             As Long = &H400 ' Zeigt nur Windows- oder Unicode-Fonts
Private Const CF_SCRIPTSONLY          As Long = &H400 ' Listet nur Windows- oder Unicode-Fonts
Private Const CF_NOOEMFONTS           As Long = &H800 ' Zeigt keine OEM Fonts
Private Const CF_NOVECTORFONTS        As Long = &H800 ' keine Vector-Fonts anzeigen
Private Const CF_NOSIMULATIONS        As Long = &H1000 ' Kein Beispiel (Vorschau) anzeigen
Private Const CF_LIMITSIZE            As Long = &H2000 ' Erlaubt nur Schriftgrößen im Bereich "nSizeMin" und "nSizeMax"
Private Const CF_FIXEDPITCHONLY       As Long = &H4000 ' Listet nur Fixed-Pitch Fonts
Private Const CF_WYSIWYG              As Long = &H8000  ' Listet nur Fonts, die Drucker- und Bildschirm-Fonts gleichzeitig sind' (muss benutzt werden mit CF_BOTH und CF_SCALABLEONLY)
Private Const CF_FORCEFONTEXIST       As Long = &H10000 ' Verweigert die Eingabe nicht aufgeführter Fonts
Private Const CF_SCALABLEONLY         As Long = &H20000 ' Listet nur skalierbare Fonts
Private Const CF_TTONLY               As Long = &H40000 ' Listet nur TrueType-Schriftarten
Private Const CF_NOFACESEL            As Long = &H80000 ' Kein Standard Facenamen selektieren
Private Const CF_NOSTYLESEL           As Long = &H100000 ' kein Standard-Stil setzen
Private Const CF_NOSIZESEL            As Long = &H200000 ' keine Standardgröße setzen
Private Const CF_SELECTSCRIPT         As Long = &H400000 ' Listet nur Script-Fonts
Private Const CF_NOSCRIPTSEL          As Long = &H800000 ' Kein Standard Script selektieren
Private Const CF_NOVERTFONTS          As Long = &H1000000 ' keine vertikal ausgerichteten Fonts anzeigen

' nFontType Konstanten
Private Const BOLD_FONTTYPE = &H100       ' fett
Private Const ITALIC_FONTTYPE = &H200     ' kursiv
Private Const REGULAR_FONTTYPE = &H400    ' reguläre Fonts
Private Const SCREEN_FONTTYPE = &H2000    ' Bildschirm-Fonts
Private Const PRINTER_FONTTYPE = &H4000   ' Drucker-Fonts
Private Const SIMULATED_FONTTYPE = &H8000 ' Fonts, die in der Dialogbox simuliert werden können

' lfWeight Konstanten
Private Const FW_DONTCARE = 0     ' Standard
Private Const FW_THIN = 100       ' super dünn
Private Const FW_EXTRALIGHT = 200 ' extra dünn
Private Const FW_LIGHT = 300      ' dünn
Private Const FW_NORMAL = 400     ' normal
Private Const FW_MEDIUM = 500     ' mittel
Private Const FW_SEMIBOLD = 600   ' etwas dicker
Private Const FW_BOLD = 700       ' fett
Private Const FW_EXTRABOLD = 800  ' extra fett
Private Const FW_HEAVY = 900      ' super fett

' lfCharSet Konstanten 'Skript
Private Const ANSI_CHARSET = 0          ' Ansi Zeichensatz
Private Const DEFAULT_CHARSET = 1       ' Standard
Private Const SYMBOL_CHARSET = 2        ' Symbolisch
Private Const MAC_CHARSET = 77          ' Mac (Win 9x)
Private Const SHIFTJIS_CHARSET = 128    ' ShiftJis
Private Const HANGEUL_CHARSET = 129     ' Handgeul
Private Const JOHAB_CHARSET = 130       ' Johab (Win 9x)
Private Const GB2312_CHARSET = 134      ' Englisch
Private Const CHINESEBIG5_CHARSET = 136 ' Chinesisch
Private Const GREEK_CHARSET = 161       ' Griechisch (Win 9x)
Private Const TURKISH_CHARSET = 162     ' Türkisch (Win 9x)
Private Const HEBREW_CHARSET = 177      ' Hebräisch (NT/2000)
Private Const ARABIC_CHARSET = 178      ' Arabisch (NT/2000)
Private Const BALTIC_CHARSET = 186      ' Baltisch (Win 9x)
Private Const RUSSIAN_CHARSET = 204     ' Russisch (Win 9x)
Private Const EASTEUROPE_CHARSET = 238  ' Osteuropäisch (Win 9x)
Private Const THAI_CHARSET = 222        ' Thailändisch (NT/2000)
Private Const OEM_CHARSET = 255         ' OEM

' lfOutPrecision Konstanten
Private Const OUT_DEFAULT_PRECIS = 0 ' Verwendet Standard-Fonts
Private Const OUT_DEVICE_PRECIS = 5  ' Verwendet eine Device-Font wenn mehrere Dateien mit dem selben Namen existieren
Private Const OUT_OUTLINE_PRECIS = 8 ' (NT/2000) Verwendet eine TureType Schriftart oder Outline-Based Fonts.
Private Const OUT_RASTER_PRECIS = 6  ' Verwendet ein Raster-Font, wenn mehrere Dateien mit dem selben Namen existieren
Private Const OUT_STRING_PRECIS = 1  ' Raster-Font für die Enumeration
Private Const OUT_STROKE_PRECIS = 3  ' (Win 9x) Vector-Font für die Enumeration' (NT/2000) TrueType, Outline-Based oder VectorFont für die Enumeration
Private Const OUT_TT_ONLY_PRECIS = 7 ' Verwendet eine TrueType-Schriftart
Private Const OUT_TT_PRECIS = 4      ' Verwendet eine TrueType-Font, wenn mehrere' Dateien mit dem selben Namen existieren

' lfClipPrecision Konstanten
Private Const CLIP_DEFAULT_PRECIS = 0 ' Standard
Private Const CLIP_EMBEDDED = 128 ' wird benutzt für eingebettete schreibgeschützte Fonts
Private Const CLIP_LH_ANGLES = 16 ' die Richtung von irgendwelchen Rotationen wird vom Koordinatensystem festgelegt (gegen den Uhrzeigersinn)
Private Const CLIP_STROKE_PRECIS = 2 ' Raster-Vector oder TrueType-Fonts für die Enumeration

' lfQuality Konstanten
Private Const ANTIALIASED_QUALITY = 4    ' (Win 9x, NT 4.0, 2000) Schriftart wird wenn möglich immer mit Anitialisierung gezeichnet
Private Const DEFAULT_QUALITY = 0        ' Standard Qualität
Private Const DRAFT_QUALITY = 1          ' entspricht "Schnelldruck" bei Druckausgaben
Private Const NONANTIALIASED_QUALITY = 3 ' (Win 9x, NT 4.0, 2000) Schriftart wird nicht mit Anitialisierung gezeichnet
Private Const PROOF_QUALITY = 2          ' entspricht "Schöndruck" bei Druckausgaben

' lfPitchAndFamily Konstanten
Private Const DEFAULT_PITCH = 0  ' Standard
Private Const FIXED_PITCH = 1    ' Fest
Private Const VARIABLE_PITCH = 2 ' Variabel
Private Const FF_DECORATIVE = 80 ' Dekoriert
Private Const FF_DONTCARE = 0    ' Egal
Private Const FF_MODERN = 48     ' Modern
Private Const FF_ROMAN = 16      ' Roman
Private Const FF_SCRIPT = 64     ' Script
Private Const FF_SWISS = 32      ' Swiss

Private Const LOGPIXELSY As Long = 90
Private mDPI As Long
Private mPT2Px As Long
Private mFlags As Long
Private mColor As Long
Private mFont As LOGFONT
Private mMinSize As Long
Private mMaxSize As Long
Private mTag As Variant

Private Sub Class_Initialize()
  mDPI = GetDeviceCaps(GetDC(0), LOGPIXELSY)
  mPT2Px = 72
  Reset
End Sub

'Public Overrides Sub Reset()
Public Sub Reset()
Dim StdName As String: StdName = "Ms Sans Serif"
Dim StdF As New StdFont:
  StdF.Name = StdName
' Setzt alle Optionen im Dialogfeld auf die Standardwerte zurück.
  mFlags = mFlags Or CF_BOTH
  mFlags = mFlags Or CF_EFFECTS
  mFlags = mFlags Or CF_INITTOLOGFONTSTRUCT
  mFont.lfHeight = S2H(8) '10 + 2 / 3
  mFont.lfFaceName = StrConv(StdF.Name, vbFromUnicode) & vbNullChar '& Space$(LF_FACESIZE - 2 - Len(StdName)) & vbNullChar
End Sub

Private Function S2H(aSize As Long) As Long
  S2H = -MulDiv(aSize, mDPI, mPT2Px)
End Function
Private Function H2S(aHeight As Long) As Long
  H2S = -MulDiv(aHeight, mPT2Px, mDPI)
End Function
'Public Property AllowScriptChange() As Boolean
Public Property Get AllowScriptChange() As Boolean
' Ruft einen Wert ab, der angibt, ob der Benutzer den im Kombinationsfeld Script
' angegebenen Zeichensatz ändern kann, so dass anstelle des aktuellen ein anderer
' Zeichensatz angezeigt wird, oder legt diesen fest.
  AllowScriptChange = Not (mFlags And CF_NOSCRIPTSEL)
End Property
Public Property Let AllowScriptChange(BolVal As Boolean)
  mFlags = mFlags Or CF_NOSCRIPTSEL
  If BolVal Then
    mFlags = mFlags Xor CF_NOSCRIPTSEL
  End If
End Property

'Public Property AllowSimulations() As Boolean
Public Property Get AllowSimulations() As Boolean
' Ruft einen Wert ab, der angibt, ob GDI-Schriftartsimulationen
' (Graphics Device Interface) für das Dialogfeld zugelassen sind, oder legt diesen fest.
'mit einfacheren Worten: Fontvorschau mit kurzem Schriftbeispiel
  AllowSimulations = Not (mFlags And CF_NOSIMULATIONS)
End Property
Public Property Let AllowSimulations(BolVal As Boolean)
  mFlags = mFlags Or CF_NOSIMULATIONS
  If BolVal Then
    mFlags = mFlags Xor CF_NOSIMULATIONS
  End If
End Property

'Public Property AllowVectorFonts() As Boolean
Public Property Get AllowVectorFonts() As Boolean
' Ruft einen Wert ab, der angibt, ob Vektorschriftarten im Dialogfeld ausgewählt werden
' können, oder legt diesen fest.
  AllowVectorFonts = Not (mFlags And CF_NOVECTORFONTS)
End Property
Public Property Let AllowVectorFonts(BolVal As Boolean)
  mFlags = mFlags Or CF_NOVECTORFONTS
  If BolVal Then
    mFlags = mFlags Xor CF_NOVECTORFONTS
  End If
End Property

'Public Property AllowVerticalFonts() As Boolean
Public Property Get AllowVerticalFonts() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld vertikale und horizontale Schriftarten
' oder nur horizontale Schriftarten angezeigt werden, oder legt diesen fest.
  AllowVerticalFonts = Not (mFlags And CF_NOVERTFONTS)
End Property
Public Property Let AllowVerticalFonts(BolVal As Boolean)
  mFlags = mFlags Or CF_NOVERTFONTS
  If BolVal Then
    mFlags = mFlags Xor CF_NOVERTFONTS
  End If
End Property

'Public Property Color() As System.Drawing.Color
Public Property Get Color() As Long
' Ruft die ausgewählte Schriftfarbe ab oder legt diese fest.
  Color = mColor
End Property
Public Property Let Color(LngVal As Long)
  mColor = LngVal
End Property

'Public Property FixedPitchOnly() As Boolean
Public Property Get FixedPitchOnly() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld nur Schriftarten mit fester
' Schriftbreite ausgewählt werden können, oder legt diesen fest.
   FixedPitchOnly = (mFlags And CF_FIXEDPITCHONLY)
End Property
Public Property Let FixedPitchOnly(BolVal As Boolean)
  mFlags = mFlags Or CF_FIXEDPITCHONLY
  If Not BolVal Then
    mFlags = mFlags Xor CF_FIXEDPITCHONLY
  End If
End Property

'Public Property Font() As System.Drawing.Font
Public Property Get Font() As Font 'Boolean
  Set Font = New StdFont
  With Font
    .Bold = mFont.lfWidth
    .Charset = mFont.lfCharSet
    .Italic = mFont.lfItalic
    .Name = StrConv(mFont.lfFaceName, vbUnicode)
    .Size = H2S(mFont.lfHeight)
    .Strikethrough = mFont.lfStrikeOut
    .Underline = mFont.lfUnderline
    .Weight = mFont.lfWeight
  End With
End Property
Public Property Set Font(FntVal As Font)
  With mFont
    .lfWidth = FntVal.Bold
    .lfCharSet = FntVal.Charset
    .lfItalic = FntVal.Italic
    .lfFaceName = StrConv(FntVal.Name, vbFromUnicode) & vbNullChar
    .lfHeight = S2H(FntVal.Size)
    .lfStrikeOut = FntVal.Strikethrough
    .lfUnderline = FntVal.Underline
    .lfWeight = FntVal.Weight
  End With
End Property

'Public Property FontMustExist() As Boolean
Public Property Get FontMustExist() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld eine Fehlerbedingung angegeben wird,
' wenn der Benutzer eine nicht vorhandene Schriftart oder ein nicht vorhandenes Format
' auswählt, oder legt diesen fest.
   FontMustExist = (mFlags And CF_FORCEFONTEXIST)
End Property
Public Property Let FontMustExist(BolVal As Boolean)
  mFlags = mFlags Or CF_FORCEFONTEXIST
  If Not BolVal Then
    mFlags = mFlags Xor CF_FORCEFONTEXIST
  End If
End Property

'Public Property MaxSize() As Integer
Public Property Get MaxSize() As Long
' Ruft den größten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.
  MaxSize = mMaxSize
End Property
Public Property Let MaxSize(LngVal As Long)
  mMaxSize = LngVal
End Property

'Public Property MinSize() As Integer
Public Property Get MinSize() As Long
' Ruft den kleinsten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.
  MinSize = mMinSize
End Property
Public Property Let MinSize(LngVal As Long)
  mMinSize = LngVal
End Property

'Public Property ScriptsOnly() As Boolean
Public Property Get ScriptsOnly() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld Schriftarten für alle
' Nicht-OEM-Zeichensätze und Symbolzeichensätze sowie der ANSI-Zeichensatz
' ausgewählt werden können, oder legt diesen fest.
   ScriptsOnly = (mFlags And CF_SELECTSCRIPT)
End Property
Public Property Let ScriptsOnly(BolVal As Boolean)
  mFlags = mFlags Or CF_SELECTSCRIPT
  If Not BolVal Then
    mFlags = mFlags Xor CF_SELECTSCRIPT
  End If
End Property

'Public Property ShowApply() As Boolean
Public Property Get ShowApply() As Boolean
' Ruft einen Wert ab, der angibt, ob das Dialogfeld die Schaltfläche Apply enthält,
' oder legt diesen fest.
   ShowApply = (mFlags And CF_APPLY)
End Property
Public Property Let ShowApply(BolVal As Boolean)
  mFlags = mFlags Or CF_APPLY
  If Not BolVal Then
    mFlags = mFlags Xor CF_APPLY
  End If
End Property

'Public Property ShowColor() As Boolean
Public Property Get ShowColor() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld die Farbauswahl angezeigt wird, oder legt diesen fest.

End Property
Public Property Let ShowColor(BolVal As Boolean)

End Property

'Public Property ShowEffects() As Boolean
Public Property Get ShowEffects() As Boolean
' Ruft einen Wert ab, der angibt, ob das Dialogfeld Steuerelemente enthält, mit denen der
' Benutzer Optionen zum Durchstreichen, Unterstreichen und für die Textfarbe angeben kann
' oder legt diesen fest.
   ShowEffects = (mFlags And CF_EFFECTS)
End Property
Public Property Let ShowEffects(BolVal As Boolean)
  mFlags = mFlags Or CF_EFFECTS
  If Not BolVal Then
    mFlags = mFlags Xor CF_EFFECTS
  End If
End Property

'Public Property ShowHelp() As Boolean
Public Property Get ShowHelp() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld die Hilfeschaltfläche angezeigt wird, oder legt diesen fest.
   ShowHelp = (mFlags And CF_SHOWHELP)
End Property
Public Property Let ShowHelp(BolVal As Boolean)
  mFlags = mFlags Or CF_SHOWHELP
  If Not BolVal Then
    mFlags = mFlags Xor CF_SHOWHELP
  End If
End Property

'Protected ReadOnly Property Options() As Integer
Public Property Get Options() As Long
  Options = mFlags
End Property


'wie schon hinlänglich bekannt, fügt einfach irgendwas hinzu
'wird intern nicht verwendet, kann extern verwendet werden
'Public Property Get Tag() As Object
Public Property Get Tag() As Variant
  Tag = mTag
End Property
Public Property Let Tag(VarVal As Variant)
  mTag = VarVal
End Property

Public Function ToString() As String
  ToString = "System.Windows.Forms.FolderBrowserDialog"
End Function

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult
Dim CF As ChooseFont
Dim hhwndOwner As Long

TryE: On Error GoTo Catch
  If IsMissing(Frm) Then
    hhwndOwner = GetActiveWindow
  Else
    hhwndOwner = Frm.hwnd
  End If
  
  With CF
    .lStructSize = Len(CF)
    .hwndOwner = hhwndOwner
    .hInstance = App.hInstance
    .hdc = Printer.hdc
    .flags = mFlags
    .iPointSize = 10
    .lpLogFont = VarPtr(mFont)
    .rgbColors = mColor
    .nSizeMax = mMaxSize
    .nSizeMin = mMinSize
  End With
  ShowDialog = ChooseFontA(CF)
  mColor = CF.rgbColors
  mMinSize = CF.nSizeMin
  mMaxSize = CF.nSizeMax
  Exit Function
Catch:
  ShowDialog = vbAbort
End Function

