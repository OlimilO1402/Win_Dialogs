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
Private Const CF_EFFECTS              As Long = &H100&    ' Erlaubt Font-Eigenschaften wie' Unterstreichen, Farbe und Durchgestrichen
Private Const CF_APPLY                As Long = &H200&    ' Zeigt einen "Übernehmen" Button an
Private Const CF_SCRIPTSONLY          As Long = &H400&    ' Listet nur Windows- oder Unicode-Fonts
Private Const CF_NOVECTORFONTS        As Long = &H800&    ' keine Vector-Fonts anzeigen
Private Const CF_NOSIMULATIONS        As Long = &H1000&   ' Kein Beispiel (Vorschau) anzeigen
Private Const CF_FIXEDPITCHONLY       As Long = &H4000&   ' Listet nur Fixed-Pitch Fonts
Private Const CF_FORCEFONTEXIST       As Long = &H10000   ' Verweigert die Eingabe nicht aufgeführter Fonts
Private Const CF_TTONLY               As Long = &H40000   ' Listet nur TrueType-Schriftarten
Private Const CF_SELECTSCRIPT         As Long = &H400000  ' Listet nur Script-Fonts
Private Const CF_NOVERTFONTS          As Long = &H1000000 ' keine vertikal ausgerichteten Fonts anzeigen

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
Private Type LOGFONT 'W
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
    hWndOwner      As LongPtr '8  caller's window handle
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


Private Const defaultMinSize As Long = 0

Private Const defaultMaxSize  As Long = 0

Private mOptions As Long

Private mFont As Font

Private mColor As Long 'Color

Private mMinSize As Long

Private mMaxSize As Long

Private mShowColor As Boolean

Private usingDefaultIndirectColor As Boolean

'/// <summary>Ruft einen Wert ab, der angibt, ob GDI-Schriftartsimulationen (Graphics Device Interface) für das Dialogfeld zugelassen sind, oder legt diesen fest.</summary>
'/// <returns>true, wenn Schriftartsimulationen zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(true)]
'[SRDescription("FnDallowSimulationsDescr")]
Public Property Get AllowSimulations() As Boolean
    AllowSimulations = Not GetOption(CF_NOSIMULATIONS)
End Property
Public Property Let AllowSimulations(ByVal value As Boolean)
    SetOption CF_NOSIMULATIONS, Not value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob Vektorschriftarten im Dialogfeld ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn Vektorschriftarten zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(true)]
'[SRDescription("FnDallowVectorFontsDescr")]
Public Property Get AllowVectorFonts() As Boolean
    AllowVectorFonts = Not GetOption(CF_NOVECTORFONTS)
End Property
Public Property Let AllowVectorFonts(ByVal value As Boolean)
    SetOption CF_NOVECTORFONTS, Not value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld vertikale und horizontale Schriftarten oder nur horizontale Schriftarten angezeigt werden, oder legt diesen fest.</summary>
'/// <returns>true, wenn vertikale und horizontale Schriftarten zulässig sind, andernfalls false.Der Standardwert ist true.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(true)]
'[SRDescription("FnDallowVerticalFontsDescr")]
Public Property Get AllowVerticalFonts() As Boolean
    AllowVerticalFonts = Not GetOption(CF_NOVERTFONTS)
End Property
Public Property Let AllowVerticalFonts(ByVal value As Boolean)
        SetOption CF_NOVERTFONTS, !value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob der Benutzer den im Kombinationsfeld Skript angegebenen Zeichensatz ändern kann, sodass anstelle des aktuellen ein anderer Zeichensatz angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn der Benutzer den im Kombinationsfeld Skript angegebenen Zeichensatz ändern kann, andernfalls false.Der Standardwert ist true.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(true)]
'[SRDescription("FnDallowScriptChangeDescr")]
Public Property Get AllowScriptChange() As Boolean
    AllowScriptChange = Not GetOption(CF_SELECTSCRIPT)
End Property
Public Property Let AllowScriptChange(ByVal value As Boolean)
    SetOption CF_SELECTSCRIPT, Not value
End Property

'/// <summary>Ruft die ausgewählte Schriftfarbe ab oder legt diese fest.</summary>
'/// <returns>Die Farbe der ausgewählten Schriftart.Der Standardwert ist <see cref="P:System.Drawing.Color.Black" />.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatData")]
'[SRDescription("FnDcolorDescr")]
'[DefaultValue(typeof(Color), "Black")]
Public Property Get Color() As Variant 'Long 'Color
    If usingDefaultIndirectColor Then
        Color = ColorTranslator.FromWin32(ColorTranslator.ToWin32(Color)) '??? WTF Wozu?
    Else
        Color = mColor
    End If
End Property
Public Property Let Color(ByVal value As Variant) 'Color
    If Not IsEmpty(value) Then
        mColor = value
        usingDefaultIndirectColor = False
    Else
        mColor = vbBlack 'SystemColors.ControlText'lets assume this is black
        usingDefaultIndirectColor = True
    End If
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld nur Schriftarten mit fester Schriftbreite ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn nur Schriftarten mit fester Schriftbreite ausgewählt werden können, andernfalls false.The default value is false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDfixedPitchOnlyDescr")]
Public Property Get FixedPitchOnly() As Boolean
    FixedPitchOnly = GetOption(CF_FIXEDPITCHONLY)
End Property
Public Property Let FixedPitchOnly(ByVal value As Boolean)
    SetOption CF_FIXEDPITCHONLY, value
End Property

'/// <summary>Ruft die ausgewählte Schriftart ab oder legt diese fest.</summary>
'/// <returns>Die ausgewählte Schriftart.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatData")]
'[SRDescription("FnDfontDescr")]
Public Property Get Font() As Font
    Dim aFont As Font: Set aFont = mFont
    If aFont Is Nothing Then
        Set Font = Control.DefaultFont
    End If
    Dim SizeInPoints As Single: SizeInPoints = Font.SizeInPoints
    If MinSize <> 0 And SizeInPoints < CSng(MinSize) Then
        Set Font = MNew.Font(Font.FontFamily, MinSize, Font.Style, GraphicsUnit.Point)
    End If
    If MaxSize <> 0 And SizeInPoints > CSng(MaxSize) Then
        Set Font = MNew.Font(Font.FontFamily, MaxSize, Font.Style, GraphicsUnit.Point)
    End If
    Set Font = aFont
End Property
Public Property Set Font(ByVal value As Font)
    Set mFont = value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld eine Fehlerbedingung angegeben wird, wenn der Benutzer eine nicht vorhandene Schriftart oder ein nicht vorhandenes Format auswählt, oder legt diesen fest.</summary>
'/// <returns>true, wenn im Dialogfeld eine Fehlerbedingung angezeigt wird, wenn der Benutzer eine nicht vorhanden Schriftart bzw. ein nicht vorhandenes Format auswählt, andernfalls false.Der Standardwert ist false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDfontMustExistDescr")]
Public Property Get FontMustExist() As Boolean
    FontMustExist = GetOption(CF_FORCEFONTEXIST)
End Property
Public Property Let FontMustExist(ByVal value As Boolean)
    SetOption CF_FORCEFONTEXIST, value
End Property

'/// <summary>Ruft den größten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.</summary>
'/// <returns>Der größte Schriftgrad, den der Benutzer auswählen kann. Der Standardwert ist 0.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatData")]
'[DefaultValue(0)]
'[SRDescription("FnDmaxSizeDescr")]
Public Property Get MaxSize() As Long
    MaxSize = mMaxSize
End Property
Public Property Let MaxSize(ByVal value As Long)
    If value < 0 Then
        value = 0
    End If
    mMaxSize = value
    If 0 < mMaxSize And mMaxSize < mMinSize Then
        mMinSize = MaxSize
    End If
End Property

'/// <summary>Ruft den kleinsten Schriftgrad ab, den der Benutzer auswählen kann, oder legt diesen fest.</summary>
'/// <returns>Der kleinste Schriftgrad, den der Benutzer auswählen kann.Der Standardwert ist 0.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatData")]
'[DefaultValue(0)]
'[SRDescription("FnDminSizeDescr")]
Public Property Get MinSize() As Long
    MinSize = mMinSize
End Property
Public Property Let MinSize(ByVal value As Long)
    If value < 0 Then
        value = 0
    End If
    mMinSize = value
    If 0 < mMaxSize And mMaxSize < mMinSize Then
        mMaxSize = MinSize
    End If
End Property

Public Property Get Options() As Long
    Options = mOptions
End Property
'/// <summary>Ruft Werte für die Initialisierung von <see cref="T:System.Windows.Forms.FontDialog" /> ab.</summary>
'/// <returns>Eine bitweise Kombination von internen Werten, die das <see cref="T:System.Windows.Forms.FontDialog" /> initialisiert.</returns>
'private  int Options => options

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld Schriftarten für alle Nicht-OEM-Zeichensätze und Symbolzeichensätze sowie der ANSI-Zeichensatz ausgewählt werden können, oder legt diesen fest.</summary>
'/// <returns>true, wenn alle Nicht-OEM-Zeichensätze und Symbolzeichensätze sowie der ANSI-Zeichensatz ausgewählt werden können, andernfalls false.The default value is false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDscriptsOnlyDescr")]
Public Property Get ScriptsOnly() As Boolean
    ScriptsOnly = GetOption(CF_SCRIPTSONLY)
End Property
Public Property Let ScriptsOnly(ByVal value As Boolean)
    SetOption CF_SCRIPTSONLY, value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob das Dialogfeld die Schaltfläche Übernehmen enthält, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld die Schaltfläche Übernehmen enthält, andernfalls false.Der Standardwert ist false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDshowApplyDescr")]
Public Property Get ShowApply() As Boolean
    ShowApply = GetOption(CF_APPLY)
End Property
Public Property Let ShowApply(ByVal value As Boolean)
    SetOption CF_APPLY, value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld die Farbauswahl angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn im Dialogfeld die Farbauswahl angezeigt wird, andernfalls false.Der Standardwert ist false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDshowColorDescr")]
Public Property Get ShowColor() As Boolean
    ShowColor = mShowColor
End Property
Public Property Let ShowColor(ByVal value As Boolean)
    mShowColor = value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob das Dialogfeld Steuerelemente enthält, mit denen der Benutzer Optionen zum Durchstreichen, Unterstreichen und für die Textfarbe angeben kann, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld Steuerelemente enthält, mit denen das Durchstreichen, Unterstreichen und die Textfarbe festgelegt werden können, andernfalls false.Der Standardwert ist true.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(true)]
'[SRDescription("FnDshowEffectsDescr")]
Public Property Get ShowEffects() As Boolean
    ShowEffects = GetOption(CF_EFFECTS)
End Property
Public Property Let ShowEffects(ByVal value As Boolean)
    SetOption CF_EFFECTS, value
End Property

'/// <summary>Ruft einen Wert ab, der angibt, ob im Dialogfeld die Hilfeschaltfläche angezeigt wird, oder legt diesen fest.</summary>
'/// <returns>true, wenn das Dialogfeld die Schaltfläche Hilfe enthält, andernfalls false.Der Standardwert ist false.</returns>
'/// <filterpriority>1</filterpriority>
'[SRCategory("CatBehavior")]
'[DefaultValue(false)]
'[SRDescription("FnDshowHelpDescr")]
Public Property Get ShowHelp() As Boolean
    ShowHelp = GetOption(CF_SHOWHELP)
End Property
Public Property Let ShowHelp(ByVal value As Boolean)
    SetOption CF_SHOWHELP, value
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

'/// <summary>Initialisiert eine neue Instanz der <see cref="T:System.Windows.Forms.FontDialog" />-Klasse.</summary>
Private Sub Class_Initialize()
    Reset
End Sub

Private Function GetOption(ByVal aOption As Long) As Boolean
    GetOption = (mOptions And aOption) <> 0
End Function

'/// <summary>Gibt für Standarddialogfelder die Hookprozedur an, die überschrieben wird, um einem Standarddialogfeld bestimmte Funktionen hinzuzufügen.</summary>
'/// <returns>Ein Wert von 0, wenn die Meldung von der Prozedur für Standarddialogfelder verarbeitet wird. Ein Wert ungleich 0, wenn die Meldung von dieser Prozedur ignoriert wird.</returns>
'/// <param name="hWnd">Das Handle für das Dialogfeldfenster. </param>
'/// <param name="msg">Die empfangene Meldung. </param>
'/// <param name="wparam">Zusätzliche Informationen zur Meldung. </param>
'/// <param name="lparam">Zusätzliche Informationen zur Meldung. </param>
'[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
'    private  override IntPtr HookProc(IntPtr hWnd, int msg, IntPtr wparam, IntPtr lparam)
'    {
'        Switch (msg)
'        {
'        Case 273:
'        {
'            if ((int)wparam != 1026)
'            {
'                break
'            }
'            NativeMethods.LOGFONT lOGFONT = new NativeMethods.LOGFONT()
'            UnsafeNativeMethods.SendMessage(new HandleRef(null, hWnd), 1025, 0, lOGFONT)
'            UpdateFont (LOGFONT)
'            int num = (int)UnsafeNativeMethods.SendDlgItemMessage(new HandleRef(null, hWnd), 1139, 327, IntPtr.Zero, IntPtr.Zero)
'            if (num != -1)
'            {
'                UpdateColor((int)UnsafeNativeMethods.SendDlgItemMessage(new HandleRef(null, hWnd), 1139, 336, (IntPtr)num, IntPtr.Zero))
'            }
'            if (NativeWindow.WndProcShouldBeDebuggable)
'            {
'                OnApply (EventArgs.Empty)
'                break
'            }
'            Try
'            {
'                OnApply (EventArgs.Empty)
'            }
'            catch (Exception t)
'            {
'                Application.OnThreadException (t)
'            }
'            break
'        }
'        Case 272:
'            if (!showColor)
'            {
'                IntPtr dlgItem = UnsafeNativeMethods.GetDlgItem(new HandleRef(null, hWnd), 1139)
'                SafeNativeMethods.ShowWindow(new HandleRef(null, dlgItem), 0)
'                dlgItem = UnsafeNativeMethods.GetDlgItem(new HandleRef(null, hWnd), 1091)
'                SafeNativeMethods.ShowWindow(new HandleRef(null, dlgItem), 0)
'            }
'            break
'        }
'        return base.HookProc(hWnd, msg, wparam, lparam)
'    }

'/// <summary>Löst das <see cref="E:System.Windows.Forms.FontDialog.Apply" />-Ereignis aus.</summary>
'/// <param name="e">Eine Instanz von <see cref="T:System.EventArgs" />, die die Daten enthält. </param>
'    private  virtual void OnApply(EventArgs e)
'    {
'        ((EventHandler)base.Events'[EventApply])?.Invoke(this, e)
'    }

'/// <summary>Setzt alle Optionen im Dialogfeld auf die Standardwerte zurück.</summary>
'/// <filterpriority>1</filterpriority>
Public Sub Reset()
    mOptions = CF_EFFECTS Or CF_SCREENFONTS ' 257
    Set mFont = Nothing
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

'/// <summary>Gibt ein Dateidialogfeld an.</summary>
'/// <returns>true, wenn das Dialogfeld erfolgreich ausgeführt wurde, andernfalls false.</returns>
'/// <param name="hWndOwner">Das Fensterhandle des Besitzerfensters für das Standarddialogfeld.</param>
'private  override bool RunDialog(IntPtr hWndOwner)
Public Function ShowDialog(ByVal hWndOwner As Long) As VbMsgBoxResult 'Boolean
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
Try: On Error GoTo Finally
'{
    Dim IntPtr_LF As LongPtr ': IntPtr_LF = IntPtr.Zero
    Dim aLOGFONT  As LOGFONT ': = new NativeMethods.LOGFONT()
    Font_ToLogFont aLOGFONT ', graphics
    IntPtr_LF = CoTaskMemAlloc(LenB(aLOGFONT)) 'Marshal.SizeOf(typeof(NativeMethods.LOGFONT)))
    'Marshal.StructureToPtr((object)lOGFONT, intPtr, fDeleteOld: false)
    RtlMoveMemory ByVal IntPtr_LF, aLOGFONT, LenB(aLOGFONT)
    Dim aCHOOSEFONT As CHOOSEFONT
    aCHOOSEFONT.lStructSize = LenB(aCHOOSEFONT) 'Marshal.SizeOf(typeof(NativeMethods.CHOOSEFONT))
    aCHOOSEFONT.hWndOwner = hWndOwner
    aCHOOSEFONT.hDC = 0 'IntPtr.Zero
    aCHOOSEFONT.lpLogFont = IntPtr_LF
    aCHOOSEFONT.flags = mOptions Or &H40 Or &H8
    If (mMinSize > 0 Or mMaxSize > 0) Then
        aCHOOSEFONT.flags = aCHOOSEFONT.flags Or &H2000
    End If
    If (mShowColor Or ShowEffects) Then
        aCHOOSEFONT.rgbColors = mColor  'ColorTranslator.ToWin32(Color)
    Else
        aCHOOSEFONT.rgbColors = vbBlack 'ColorTranslator.ToWin32(SystemColors.ControlText)
    End If
    'aCHOOSEFONT.lpfnHook = lpfnHook
    aCHOOSEFONT.hInstance = GetModuleHandle(0) 'UnsafeNativeMethods.GetModuleHandle(Null)
    aCHOOSEFONT.nSizeMin = mMinSize
    If (mMaxSize = 0) Then
        aCHOOSEFONT.nSizeMax = 2147483647 'int.MaxValue
    Else
        aCHOOSEFONT.nSizeMax = mMaxSize
    End If
    If ChooseFontW(aCHOOSEFONT) = 0 Then
        'return false
        ShowDialog = vbCancel ' False
    End If
    Dim alOGFONT2 As LONGFONT: RtlMoveMemory alOGFONT2, IntPtr_LF, LenB(alOGFONT2) 'aLOGFONT2 = (NativeMethods.LOGFONT)UnsafeNativeMethods.PtrToStructure(intPtr, typeof(NativeMethods.LOGFONT))
    If (alOGFONT2.lfFaceName <> 0 And Len(lOGFONT2.lfFaceName) > 0) Then
        aLOGFONT = alOGFONT2
        UpdateFont aLOGFONT
        UpdateColor aCHOOSEFONT.rgbColors
    End If
    ShowDialog = vbOK 'return true
Finally:
    If IntPtr_LF <> 0 Then  ' IntPtr.Zero)
        CoTaskMemFree IntPtr_LF
    End If
End Function

Private Sub SetOption(ByVal aOption As Long, ByVal value As Boolean)
    If value Then
        mOptions = mOptions Or aOption
    Else
        'options &= ~option
        mOptions = mOptions And Not aOption
    End If
End Sub

Private Property Get ShouldSerializeFont() As Boolean
    ShouldSerializeFont = Not Font_Equals(Control.DefaultFont)
End Property

'/// <summary>Ruft eine Zeichenfolge ab, die den Namen der derzeit im Dialogfeld ausgewählten Schriftart enthält.</summary>
'/// <returns>Eine Zeichenfolge, die den Namen der derzeit ausgewählten Schriftart enthält.</returns>
'/// <filterpriority>1</filterpriority>
Friend Function ToStr() As String
    'string str = base.ToString()
    'return str + ",  Font: " + Font.ToString()
    ToStr = "MBO.Windows.Forms.FontDialog, Font: " & mFont.Name
End Function

Private Sub UpdateColor(ByVal rgb As Long)
    If (ColorTranslator.ToWin32(Color) <> rgb) Then
        Color = ColorTranslator.FromOle(rgb)
        usingDefaultIndirectColor = False
    End If
End Sub

Private Sub UpdateFont(lf As LOGFONT)
    Dim dC As LongPtr: dC = GetDC(0)
Try: On Error GoTo Finally
    Dim aFont As Font ': Set aFont = Nothing
    'Try
    '    IntSecurity.UnmanagedCode.Assert()
    '    Try
            Set aFont = Font_FromLogFont(lf, dC)
    '    }
    '    Finally
    '    {
    '        CodeAccessPermission.RevertAssert()
    '    }
    '    this.Font = ControlPaint.FontInPoints(Font)
    '}
    'Finally
    '{
    '    font?.Dispose()
    '}
Finally:
    'UnsafeNativeMethods.ReleaseDC(NativeMethods.NullHandleRef, new HandleRef(null, dC))
    ReleaseDC dC
End Sub

'}
