VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'public class FontDialog : CommonDialog
'{
'/// <summary>Besitzt das <see cref="E:System.Windows.Forms.FontDialog.Apply" />-Ereignis.</summary>
'private  static readonly object EventApply = new object() 'wozu?

Private Const CF_SCREENFONTS          As Long = &H1&      ' Listet Bildschirm-Fonts
Private Const CF_SHOWHELP             As Long = &H4&      ' Zeigt den Hilfe-Button an
Private Const CF_ENABLEHOOK           As Long = &H8&      ' Aktiviert die Callback-Funktion
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&     ' Setzt die Startwerte, welche über die' LOGFONT-Struktur angegeben wurden
Private Const CF_EFFECTS              As Long = &H100&    ' Erlaubt Font-Eigenschaften wie' Unterstreichen, Farbe und Durchgestrichen
Private Const CF_APPLY                As Long = &H200&    ' Zeigt einen "Übernehmen" Button an
Private Const CF_SCRIPTSONLY          As Long = &H400&    ' Listet nur Windows- oder Unicode-Fonts
Private Const CF_NOVECTORFONTS        As Long = &H800&    ' keine Vector-Fonts anzeigen
Private Const CF_NOSIMULATIONS        As Long = &H1000&   ' Kein Beispiel (Vorschau) anzeigen
Private Const CF_LIMITSIZE            As Long = &H2000&   ' Erlaubt nur Schriftgrößen im Bereich "nSizeMin" und "nSizeMax"
Private Const CF_FIXEDPITCHONLY       As Long = &H4000&   ' Listet nur Fixed-Pitch Fonts
Private Const CF_FORCEFONTEXIST       As Long = &H10000   ' Verweigert die Eingabe nicht aufgeführter Fonts
Private Const CF_TTONLY               As Long = &H40000   ' Listet nur TrueType-Schriftarten
Private Const CF_SELECTSCRIPT         As Long = &H400000  ' Listet nur Script-Fonts
Private Const CF_NOVERTFONTS          As Long = &H1000000 ' keine vertikal ausgerichteten Fonts anzeigen

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

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

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

'Private Const LF_FACESIZE    As Long = 32
'Private Type LOGFONT 'W
'    lfHeight         As Long
'    lfWidth          As Long
'    lfEscapement     As Long
'    lfOrientation    As Long
'    lfWeight         As Long
'    lfItalic         As Byte
'    lfUnderline      As Byte
'    lfStrikeOut      As Byte
'    lfCharSet        As Byte
'    lfOutPrecision   As Byte
'    lfClipPrecision  As Byte
'    lfQuality        As Byte
'    lfPitchAndFamily As Byte
'    lfFaceName(0 To LF_FACESIZE - 1) As Integer 'String * LF_FACESIZE '(1 To LF_FACESIZE) As Byte
'End Type

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

Private Type ChooseFont
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
    MISSING_ALIGN  As Integer '2 'wozu, wird doch eigentlich automatisch eingefügt oder?
    nSizeMin       As Long    '4  minimum pt size allowed &
    nSizeMax       As Long    '4  max pt size allowed if
                              '   CF_LIMITSIZE is used
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
    Private Declare PtrSafe Function ChooseFontW Lib "comdlg32" (ByRef pChoosefont As Any) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long
    Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByRef pv As Any)
#Else
    Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal bytLen As Long)
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
    Private Declare Function ChooseFontW Lib "comdlg32" (ByRef pChoosefont As Any) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare Function CoTaskMemAlloc Lib "Ole32" (ByVal cb As Long) As LongPtr
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Sub CoTaskMemFree Lib "Ole32" (ByRef pv As Any)
#End If

Private Const defaultMinSize As Long = 0

Private Const defaultMaxSize As Long = 0

Private mOptions As Long

#If VBA7 Then
Private mFont As NewFont ' StdFont
#Else
Private mFont As StdFont ' NewFont
#End If

Private mColor As Long 'Color

Private mMinSize As Long

Private mMaxSize As Long

Private mShowColor As Boolean

Private usingDefaultIndirectColor As Boolean

''' ############################## '''   C'tor/Init/Reset   ''' ############################## '''
'/// <summary>Initialisiert eine neue Instanz der <see cref="T:System.Windows.Forms.FontDialog" />-Klasse.</summary>
Private Sub Class_Initialize()
    Set MFontDialog.FontDialog = Me
    Reset
End Sub

'/// <summary>Setzt alle Optionen im Dialogfeld auf die Standardwerte zurück.</summary>
'/// <filterpriority>1</filterpriority>
Public Sub Reset()
    mOptions = CF_EFFECTS Or CF_SCREENFONTS ' 257
    'Set mFont = Nothing
    Set mFont = New StdFont 'othing
    mFont.Name = "Arial"
    mFont.Size = 10
    'mfont.
    mColor = vbBlack 'SystemColors.ControlText
    usingDefaultIndirectColor = True
    mShowColor = False
    mMinSize = 0
    mMaxSize = 0
    SetOption CF_TTONLY, True
End Sub

Private Sub ResetFont()
    Set mFont = Nothing
End Sub


''' ############################## '''  Props Option related  ''' ############################## '''
'/// <summary>Ruft einen Wert ab, der angibt, ob GDI-Schriftartsimulationen (Graphics Device Interface) für das Dialogfeld zugelassen sind, oder legt diesen fest.</summary>
'/// <returns>true, wenn Schriftartsimulationen zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
Public Property Get AllowSimulations() As Boolean
    AllowSimulations = Not GetOption(CF_NOSIMULATIONS)
End Property
Public Property Let AllowSimulations(ByVal Value As Boolean)
    SetOption CF_NOSIMULATIONS, Not Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob Vektorschriftarten im Dialogfeld ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn Vektorschriftarten zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
Public Property Get AllowVectorFonts() As Boolean
    AllowVectorFonts = Not GetOption(CF_NOVECTORFONTS)
End Property
Public Property Let AllowVectorFonts(ByVal Value As Boolean)
    SetOption CF_NOVECTORFONTS, Not Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld vertikale und horizontale Schriftarten oder nur horizontale Schriftarten angezeigt werden, oder legt diesen fest.</summary>
'/// <returns>true, wenn vertikale und horizontale Schriftarten zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
Public Property Get AllowVerticalFonts() As Boolean
    AllowVerticalFonts = Not GetOption(CF_NOVERTFONTS)
End Property
Public Property Let AllowVerticalFonts(ByVal Value As Boolean)
    SetOption CF_NOVERTFONTS, Not Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob der Benutzer den im Kombinationsfeld Skript angegebenen Zeichensatz ändern kann, sodass anstelle des aktuellen ein anderer Zeichensatz angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn der Benutzer den im Kombinationsfeld Skript angegebenen Zeichensatz ändern kann, andernfalls false.Der Standardwert ist true.</returns>
Public Property Get AllowScriptChange() As Boolean
    AllowScriptChange = Not GetOption(CF_SELECTSCRIPT)
End Property
Public Property Let AllowScriptChange(ByVal Value As Boolean)
    SetOption CF_SELECTSCRIPT, Not Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld nur Schriftarten mit fester Schriftbreite ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn nur Schriftarten mit fester Schriftbreite ausgewählt werden können, andernfalls false.The default value is false.</returns>
Public Property Get FixedPitchOnly() As Boolean
    FixedPitchOnly = GetOption(CF_FIXEDPITCHONLY)
End Property
Public Property Let FixedPitchOnly(ByVal Value As Boolean)
    SetOption CF_FIXEDPITCHONLY, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld eine Fehlerbedingung angegeben wird, wenn der Benutzer eine nicht vorhandene Schriftart oder ein nicht vorhandenes Format auswählt, oder legt diesen fest.</summary>
'/// <returns>true, wenn im Dialogfeld eine Fehlerbedingung angezeigt wird, wenn der Benutzer eine nicht vorhanden Schriftart bzw. ein nicht vorhandenes Format auswählt, andernfalls false.Der Standardwert ist false.</returns>
Public Property Get FontMustExist() As Boolean
    FontMustExist = GetOption(CF_FORCEFONTEXIST)
End Property
Public Property Let FontMustExist(ByVal Value As Boolean)
    SetOption CF_FORCEFONTEXIST, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld Schriftarten für alle Nicht-OEM-Zeichensätze und Symbolzeichensätze sowie der ANSI-Zeichensatz ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn alle Nicht-OEM-Zeichensätze und Symbolzeichensätze sowie der ANSI-Zeichensatz ausgewählt werden können, andernfalls false.The default value is false.</returns>
Public Property Get ScriptsOnly() As Boolean
    ScriptsOnly = GetOption(CF_SCRIPTSONLY)
End Property
Public Property Let ScriptsOnly(ByVal Value As Boolean)
    SetOption CF_SCRIPTSONLY, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob das Dialogfeld die Schaltfläche Übernehmen enthält, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld die Schaltfläche Übernehmen enthält, andernfalls false.Der Standardwert ist false.</returns>
Public Property Get ShowApply() As Boolean
    ShowApply = GetOption(CF_APPLY)
End Property
Public Property Let ShowApply(ByVal Value As Boolean)
    SetOption CF_APPLY, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob das Dialogfeld Steuerelemente enthält, mit denen der Benutzer Optionen zum Durchstreichen, Unterstreichen und für die Textfarbe angeben kann, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld Steuerelemente enthält, mit denen das Durchstreichen, Unterstreichen und die Textfarbe festgelegt werden können, andernfalls false.Der Standardwert ist true.</returns>
Public Property Get ShowEffects() As Boolean
    ShowEffects = GetOption(CF_EFFECTS)
End Property
Public Property Let ShowEffects(ByVal Value As Boolean)
    SetOption CF_EFFECTS, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld die Hilfeschaltfläche angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld die Schaltfläche Hilfe enthält, andernfalls false.Der Standardwert ist false.</returns>
Public Property Get ShowHelp() As Boolean
    ShowHelp = GetOption(CF_SHOWHELP)
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    SetOption CF_SHOWHELP, Value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld die Farbauswahl angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn im Dialogfeld die Farbauswahl angezeigt wird, andernfalls false.Der Standardwert ist false.</returns>
Public Property Get ShowColor() As Boolean
    ShowColor = mShowColor
End Property
Public Property Let ShowColor(ByVal Value As Boolean)
    mShowColor = Value
End Property

'/// <summary>Ruft Werte für die Initialisierung von <see cref="T:System.Windows.Forms.FontDialog" /> ab.</summary>
'/// <returns>Eine bitweise Kombination von internen Werten, die das <see cref="T:System.Windows.Forms.FontDialog" /> initialisiert.</returns>
'private  int Options => options
Public Property Get Options() As Long
    Options = mOptions
End Property

Private Function GetOption(ByVal aOption As Long) As Boolean
    GetOption = (mOptions And aOption) <> 0
End Function

Private Sub SetOption(ByVal aOption As Long, ByVal Value As Boolean)
    If Value Then
        mOptions = mOptions Or aOption
    Else
        mOptions = mOptions And Not aOption
    End If
End Sub

''' ############################## '''  Public Properties  ''' ############################## '''
'/// <summary>Ruft die ausgewählte Schriftfarbe ab oder legt diese fest.</summary>
'/// <returns>Die Farbe der ausgewählten Schriftart.Der Standardwert ist <see cref="P:System.Drawing.Color.Black" />.</returns>
Public Property Get Color() As Variant 'Long 'Color
    If usingDefaultIndirectColor Then
        Color = vbBlack 'ColorTranslator.FromWin32(ColorTranslator.ToWin32(Color)) '??? WTF Wozu?
    Else
        Color = mColor
    End If
End Property
Public Property Let Color(ByVal Value As Variant) 'Color
    If Not IsEmpty(Value) Then
        mColor = Value
        usingDefaultIndirectColor = False
    Else
        mColor = vbBlack 'SystemColors.ControlText'lets assume this is black
        usingDefaultIndirectColor = True
    End If
End Property

'/// <summary>Ruft die ausgewählte Schriftart ab oder legt diese fest.</summary>
'/// <returns>Die ausgewählte Schriftart.</returns>
'Public Property Get Font() As Font
#If VBA7 Then
Public Property Get Font() As NewFont
#Else
Public Property Get Font() As StdFont
#End If
'    Dim aFont As Font: Set aFont = mFont
'    If aFont Is Nothing Then
'        Set Font = New StdFont ' Control.DefaultFont
'    End If
'    Dim SizeInPoints As Single: SizeInPoints = Font.SizeInPoints
'    If MinSize <> 0 And SizeInPoints < CSng(MinSize) Then
'        Set Font = MNew.Font(Font.FontFamily, MinSize, Font.Style, GraphicsUnit.Point)
'    End If
'    If MaxSize <> 0 And SizeInPoints > CSng(MaxSize) Then
'        Set Font = MNew.Font(Font.FontFamily, MaxSize, Font.Style, GraphicsUnit.Point)
'    End If
'    Set Font = aFont
    Set Font = mFont
End Property
'Public Property Set Font(ByVal Value As Font)
#If VBA7 Then
Public Property Set Font(ByVal Value As NewFont)
#Else
Public Property Set Font(ByVal Value As StdFont)
#End If
    Set mFont = Value
End Property

'/// <summary>Ruft den größten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.</summary>
'/// <returns>Der größte Schriftgrad, den der Benutzer auswählen kann. Der Standardwert ist 0.</returns>
Public Property Get MaxSize() As Long
    MaxSize = mMaxSize
End Property
Public Property Let MaxSize(ByVal Value As Long)
    If Value < 0 Then Value = 0
    mMaxSize = Value
    If 0 < mMaxSize And mMaxSize < mMinSize Then mMinSize = MaxSize
End Property
'/// <summary>Ruft den kleinsten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.</summary>
'/// <returns>Der kleinste Schriftgrad, den der Benutzer auswählen kann.Der Standardwert ist 0.</returns>
Public Property Get MinSize() As Long
    MinSize = mMinSize
End Property
Public Property Let MinSize(ByVal Value As Long)
    If Value < 0 Then Value = 0
    mMinSize = Value
    If 0 < mMaxSize And mMaxSize < mMinSize Then mMaxSize = MinSize
End Property

'/// <summary>Tritt ein, wenn der Benutzer im Dialogfeld für die Schriftart auf die Schaltfläche Übernehmen klickt.</summary>
'/// <filterpriority>1</filterpriority>
'[SRDescription("FnDapplyDescr")]
'    public event EventHandler Apply
'    {
'        Add
'        {
'            base.Events.AddHandler(EventApply, value)
'        }
'        Remove
'        {
'            base.Events.RemoveHandler(EventApply, value)
'        }
'    }

'/// <summary>Löst das <see cref="E:System.Windows.Forms.FontDialog.Apply" />-Ereignis aus.</summary>
'/// <param name="e">Eine Instanz von <see cref="T:System.EventArgs" />, die die Daten enthält. </param>
'    private  virtual void OnApply(EventArgs e)
'    {
'        ((EventHandler)base.Events'[EventApply])?.Invoke(this, e)
'    }



'/// <summary>Gibt ein Dateidialogfeld an.</summary>
'/// <returns>true, wenn das Dialogfeld erfolgreich ausgeführt wurde, andernfalls false.</returns>
'/// <param name="hWndOwner">Das Fensterhandle des Besitzerfensters für das Standarddialogfeld.</param>
'private  override bool RunDialog(IntPtr hWndOwner)
Public Function ShowDialog(ByVal hwndOwner As LongPtr) As VbMsgBoxResult 'Boolean
    'NativeMethods.WndProc lpfnHook = HookProc
    'NativeMethods.CHOOSEFONT cHOOSEFONT = new NativeMethods.CHOOSEFONT()
    'Dim dC As LongPtr ': dC = GetDC(0)
    'graphics graphics = graphics.FromHdcInternal(dC)
    'IntSecurity.ObjectFromWin32Handle.Assert()
    'Try
    '{
    '}
    'Finally
    '{
    '    CodeAccessPermission.RevertAssert()
    '    graphics.Dispose()
    '}
    'UnsafeNativeMethods.ReleaseDC(NativeMethods.NullHandleRef, new HandleRef(null, dC))
    'ReleaseDC dC
Try: On Error GoTo Catch
'{
    Dim lpfnHook As LongPtr: lpfnHook = FncPtr(AddressOf MFontDialog.FontDialog_HookProc)
    Set MFontDialog.FontDialog = Me
    Dim IntPtr_LF As LongPtr ': IntPtr_LF = IntPtr.Zero
    Dim aLogFont  As LOGFONT ': = new NativeMethods.LOGFONT()
    Font_ToLogFont aLogFont ', graphics
    IntPtr_LF = CoTaskMemAlloc(LenB(aLogFont)) 'Marshal.SizeOf(typeof(NativeMethods.LOGFONT)))
    GlobalLock IntPtr_LF
    'Marshal.StructureToPtr((object)lOGFONT, intPtr, fDeleteOld: false)
    RtlMoveMemory ByVal IntPtr_LF, aLogFont, LenB(aLogFont)
    Dim aCHOOSEFONT As ChooseFont
    aCHOOSEFONT.lStructSize = LenB(aCHOOSEFONT) 'Marshal.SizeOf(typeof(NativeMethods.CHOOSEFONT))
    aCHOOSEFONT.hwndOwner = hwndOwner
    aCHOOSEFONT.hDC = 0 'IntPtr.Zero
    aCHOOSEFONT.lpLogFont = IntPtr_LF
    aCHOOSEFONT.flags = mOptions Or CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK
    If (mMinSize > 0 Or mMaxSize > 0) Then
        aCHOOSEFONT.flags = aCHOOSEFONT.flags Or CF_LIMITSIZE
    End If
    If (mShowColor Or ShowEffects) Then
        aCHOOSEFONT.rgbColors = mColor  'ColorTranslator.ToWin32(Color)
    Else
        aCHOOSEFONT.rgbColors = vbBlack 'ColorTranslator.ToWin32(SystemColors.ControlText)
    End If
    aCHOOSEFONT.lpfnHook = lpfnHook
    aCHOOSEFONT.hInstance = GetModuleHandleW(0) 'UnsafeNativeMethods.GetModuleHandle(Null)
    aCHOOSEFONT.nSizeMin = mMinSize
    If (mMaxSize = 0) Then
        aCHOOSEFONT.nSizeMax = 2147483647 'int.MaxValue
    Else
        aCHOOSEFONT.nSizeMax = mMaxSize
    End If
    Dim pCF As LongPtr: pCF = CoTaskMemAlloc(LenB(aCHOOSEFONT))
    GlobalLock pCF
    RtlMoveMemory ByVal pCF, aCHOOSEFONT, LenB(aCHOOSEFONT)
    If ChooseFontW(ByVal pCF) = 0 Then
        'return false
        ShowDialog = vbCancel ' False
        CheckMessAllErrors
        Exit Function
    End If
    Dim alOGFONT2 As LOGFONT: RtlMoveMemory alOGFONT2, ByVal IntPtr_LF, LenB(alOGFONT2) 'aLOGFONT2 = (NativeMethods.LOGFONT)UnsafeNativeMethods.PtrToStructure(intPtr, typeof(NativeMethods.LOGFONT))
    'If (alOGFONT2.lfFaceName <> 0 And Len(alOGFONT2.lfFaceName) > 0) Then
        aLogFont = alOGFONT2
        UpdateFont aLogFont
        
        RtlMoveMemory aCHOOSEFONT, ByVal pCF, LenB(aCHOOSEFONT)
        UpdateColor aCHOOSEFONT.rgbColors
        GlobalUnlock pCF
        CoTaskMemFree ByVal pCF
    'End If
    ShowDialog = vbOK 'return true
    GoTo Finally
Catch:
    ShowDialog = vbCancel
    'ShowDialog = vbAbort
    CheckMessAllErrors
Finally:
    If IntPtr_LF <> 0 Then  ' IntPtr.Zero)
        GlobalUnlock ByVal IntPtr_LF
        CoTaskMemFree ByVal IntPtr_LF
    End If
    If pCF <> 0 Then
        GlobalUnlock ByVal pCF
        CoTaskMemFree ByVal pCF
    End If
End Function

Private Property Get ShouldSerializeFont() As Boolean
    'ShouldSerializeFont = Not Font_Equals(Control.DefaultFont)
End Property

'/// <summary>Ruft eine Zeichenfolge ab, die den Namen der derzeit im Dialogfeld ausgewählten Schriftart enthält.</summary>
'/// <returns>Eine Zeichenfolge, die den Namen der derzeit ausgewählten Schriftart enthält.</returns>
'/// <filterpriority>1</filterpriority>
Friend Function ToStr() As String
    'string str = base.ToString()
    'return str + ",  Font: " + Font.ToString()
    ToStr = "MBO.Windows.Forms.FontDialog, Font: " & mFont.Name
End Function

Friend Sub UpdateColor(ByVal rgb As Long)
    'If (ColorTranslator.ToWin32(mColor) <> rgb) Then
        mColor = rgb 'ColorTranslator.FromOle(rgb)
        'usingDefaultIndirectColor = False
    'End If
End Sub

Friend Sub UpdateFont(lf As LOGFONT)
    Font_FromLogFont lf
'    Dim dC As LongPtr: dC = GetDC(0)
'Try: On Error GoTo Finally
'    Dim aFont As Font ': Set aFont = Nothing
'    'Try
'    '    IntSecurity.UnmanagedCode.Assert()
'    '    Try
'            Set aFont = Font_FromLogFont(lf, dC)
'    '    }
'    '    Finally
'    '    {
'    '        CodeAccessPermission.RevertAssert()
'    '    }
'    '    this.Font = ControlPaint.FontInPoints(Font)
'    '}
'    'Finally
'    '{
'    '    font?.Dispose()
'    '}
'Finally:
'    'UnsafeNativeMethods.ReleaseDC(NativeMethods.NullHandleRef, new HandleRef(null, dC))
'    ReleaseDC dC
End Sub

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

Private Sub Font_ToLogFont(aLogFont As LOGFONT) ', graphics
    With aLogFont
        .lfCharSet = mFont.Charset
        '.lfClipPrecision = 0
        '.lfEscapement = 0
        '.lfFaceName = 0
        .lfHeight = S2H(mFont.Size)
        .lfItalic = mFont.Italic
        '.lfOrientation = 0
        '.lfOutPrecision = 0
        '.lfPitchAndFamily = 0
        '.lfQuality = 0
        .lfStrikeOut = mFont.Strikethrough
        .lfUnderline = mFont.Underline
        .lfWeight = mFont.Weight
        .lfWidth = mFont.Bold
    End With
    Dim Value As String: Value = mFont.Name
    Dim lb As Long: lb = 2 * Min(LF_FACESIZE, Len(Value))
    RtlZeroMemory ByVal VarPtr(aLogFont.lfFaceName(0)), LF_FACESIZE * 2
    RtlMoveMemory ByVal VarPtr(aLogFont.lfFaceName(0)), ByVal StrPtr(Value), lb
End Sub

'Private Function Font_FromLogFont(lf As LOGFONT, dC) As Font
Private Sub Font_FromLogFont(lf As LOGFONT) ', dC) 'As StdFont
    'Set Font_FromLogFont = New StdFont
    With mFont 'Font_FromLogFont
        .Bold = lf.lfWidth
        .Charset = lf.lfCharSet
        .Italic = lf.lfItalic
        .Size = H2S(lf.lfHeight)
        .Strikethrough = lf.lfStrikeOut
        .Underline = lf.lfUnderline
        .Weight = lf.lfWeight
    End With
    Dim FontName As String: FontName = String(LF_FACESIZE, vbNullChar)
    RtlMoveMemory ByVal StrPtr(FontName), ByVal VarPtr(lf.lfFaceName(0)), LF_FACESIZE * 2
    mFont.Name = Trim0(FontName)
End Sub

Private Function Trim0(ByVal s As String) As String
    Trim0 = VBA.Strings.Trim$(Left(s, lstrlenW(ByVal StrPtr(s))))
End Function

Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Private Function S2H(aSize As Long) As Long
    Dim mDPI As Long: mDPI = GetDeviceCaps(GetDC(0), LOGPIXELSY)
    Dim mPT2Px As Long: mPT2Px = 72
    S2H = -MulDiv(aSize, mDPI, mPT2Px)
End Function
Private Function H2S(aHeight As Long) As Long
    Dim mDPI As Long: mDPI = GetDeviceCaps(GetDC(0), LOGPIXELSY)
    Dim mPT2Px As Long: mPT2Px = 72
    H2S = -MulDiv(aHeight, mPT2Px, mDPI)
End Function

Private Sub CheckMessAllErrors()
    CheckMessVBError
    CheckMessDllError
    CheckMessCommDlgError
End Sub
Private Sub CheckMessVBError()
    Dim E As Long: E = Err.Number
    If E = 0 Then Exit Sub
    Select Case E
    Case Else: MsgBox "Error: " & E & vbCrLf & Err.Description
    End Select
End Sub
Private Sub CheckMessDllError()
    Dim E As Long: E = Err.LastDllError
    If E = 0 Then Exit Sub
    Select Case E
    Case Else: MsgBox "Error: " & E & vbCrLf & Err.Description
    End Select
End Sub
Private Sub CheckMessCommDlgError()
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

