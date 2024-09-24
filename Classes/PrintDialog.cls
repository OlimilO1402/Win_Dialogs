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
#Const IsUnicode = 1 ' sorry no unicode support until now

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Public Enum PrintAction
    PAPrintToFile    ' Die Druckausgabe erfolgt in eine Datei.
    PAPrintToPreview ' Die Druckausgabe wird als Seitenansicht dargestellt.
    PAPrintToPrinter ' Die Druckausgabe erfolgt auf den Drucker.
End Enum

Public Enum PrintRange
    AllPages = 0           ' Alle Seiten werden gedruckt.
    Selection = 1          ' Die ausgewählten Seiten werden gedruckt.
    SomePages = 2          ' Die Seiten zwischen <see cref="P:System.Drawing.Printing.PrinterSettings.FromPage" /> und <see cref="P:System.Drawing.Printing.PrinterSettings.ToPage" /> werden gedruckt.
    CurrentPage = &H400000 ' =4194304 ' Die derzeit angezeigte Seite wird gedruckt.
End Enum

Public Enum PrinterUnit     'Gibt verschiedene Maßeinheiten an, die zum Drucken verwendet werden.
    Display                 ' Die Standardeinheit (0,01 Zoll).
    ThousandthsOfAnInch     ' 1/1000 Zoll (0,001 Zoll).
    HundredthsOfAMillimeter ' 1/100 Millimeter (0,01 mm).
    TenthsOfAMillimeter     ' 1/10 Millimeter (0,1 mm).
End Enum

Public Enum PaperSourceKind ' Standardpapierzufuhren.
    Upper = 1          ' Das obere Fach eines Druckers (oder das Standardfach bei einem Drucker mit nur einem Fach).
    Lower = 2          ' Das untere Fach eines Druckers.
    Middle = 3         ' Das mittlere Fach eines Druckers.
    Manual = 4         ' Manuell zugeführtes Papier.
    Envelope = 5       ' Ein Briefumschlag.
    ManualFeed = 6     ' Manuell zugeführter Briefumschlag.
    AutomaticFeed = 7  ' Automatischer Papiereinzug.
    TractorFeed = 8    ' Ein Traktoreinzug.
    SmallFormat = 9    ' Kleinformatiges Papier.
    LargeFormat = 10   ' Großformatiges Papier.
    LargeCapacity = 11 ' Das Druckerfach mit großer Kapazität.
    Cassette = 14      ' Eine Papierkassette.
    FormSource = 15    ' Das Standardzufuhrfach des Druckers.
    Custom = 257       ' Druckerspezifische Papierzufuhr.
End Enum

Public Enum Duplex ' Gibt die Duplexeinstellungen des Druckers an.
    Default = -1   ' Die Standardduplexeinstellungen für den Drucker.
    Simplex = 1    ' Einseitiger Druck.
    Horizontal = 3 ' Beidseitiger horizontaler Druck.
    Vertical = 2   ' Beidseitiger vertikaler Druck.
End Enum

'Public Enum PrinterConstants
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
'End Enum

Private Const maxPage As Long = 9999

' PRINTDLG_TYPE flags-Konstanten
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

' DEVMODE dmFields-Konstanten
Private Const DM_ORIENTATION                As Long = &H1&       ' dmOrientation wird gefüllt/gelesen
Private Const DM_PAPERSIZE                  As Long = &H2&       ' dmPaperSize wird gefüllt/gelesen
Private Const DM_PAPERLENGTH                As Long = &H4&       ' dmPaperLenght wird gefüllt/gelesen
Private Const DM_PAPERWIDTH                 As Long = &H8&       ' dmPaperWidth wird gefüllt/gelesen
Private Const DM_SCALE                      As Long = &H10&      ' dmScale wird gefüllt/gelesen
Private Const DM_COPIES                     As Long = &H100&     ' dmCopies wird gefüllt/gelesen
Private Const DM_DEFAULTSOURCE              As Long = &H200&     ' dmDefaultSource wird gefüllt/gelesen
Private Const DM_PRINTQUALITY               As Long = &H400&     ' dmPrintQuality wird gefüllt/gelesen
Private Const DM_COLOR                      As Long = &H800&     ' dmColor wird gefüllt/gelesen
Private Const DM_DUPLEX                     As Long = &H1000&    ' dmDuplex wird gefüllt/gelesen
Private Const DM_YRESOLUTION                As Long = &H2000&    ' dmYResolution wird gefüllt/gelesen
Private Const DM_TTOPTION                   As Long = &H4000&    ' dmTTOption wird gefüllt/gelesen
Private Const DM_COLLATE                    As Long = &H8000&    ' dmCollate wird gefüllt/gelesen
Private Const DM_FORMNAME                   As Long = &H10000    ' dmFormName wird gefüllt/gelesen
Private Const DM_LOGPIXELS                  As Long = &H20000    ' dmLogPixels wird gefüllt/gelesen
Private Const DM_BITSPERPEL                 As Long = &H40000    ' dmBitsPerPixel wird gefüllt/gelesen
Private Const DM_PELSWIDTH                  As Long = &H80000    ' dmPelsWidth wird gefüllt/gelesen
Private Const DM_PELSHEIGHT                 As Long = &H100000   ' dmPelsHeight wird gefüllt/gelesen
Private Const DM_DISPLAYFLAGS               As Long = &H200000   ' dmDisplayFlags wird gefüllt/gelesen
Private Const DM_DISPLAYFREQUENCY           As Long = &H400000   ' dmDisplayFrequency wird gefüllt/gelesen
Private Const DM_ICMMETHOD                  As Long = &H800000   ' dmICMMethod wird gefüllt/gelesen
Private Const DM_ICMINTENT                  As Long = &H1000000  ' dmICMIntent wird gefüllt/gelesen
Private Const DM_MEDIATYPE                  As Long = &H2000000  ' dmMediaType wird gefüllt/gelesen
Private Const DM_DITHERTYPE                 As Long = &H4000000  ' dmDitherType wird gefüllt/gelesen
Private Const DM_PANNINGWIDTH               As Long = &H20000000 ' dmPanningWidth wird gefüllt/gelesen
Private Const DM_PANNINGHEIGHT              As Long = &H40000000 ' dmPanningHeight wird gefüllt/gelesen

' DEVMODE dmOrientation-Konstanten
Private Const DMORIENT_PORTRAIT             As Long = 1 ' Portrait  wurde gewählt (vertikal Drucken)
Private Const DMORIENT_LANDSCAPE            As Long = 2 ' Landscape wurde gewählt (horizontal Drucken)

' DEVMODE dmPaperSize-Konstanten
Private Const DMPAPER_LETTER                  As Long = 1  ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_FIRST                   As Long = 1  ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_LETTERSMALL             As Long = 2  ' Blatt klein (8,5 x 11 Inch)
Private Const DMPAPER_TABLOID                 As Long = 3  ' Tabloid Blatt (11 x 17 Inch)
Private Const DMPAPER_LEDGER                  As Long = 4  ' Ledger (17x11 Inch)
Private Const DMPAPER_LEGAL                   As Long = 5  ' Blatt 8,5 x 14 Inch
Private Const DMPAPER_STATEMENT               As Long = 6  ' Statement Blatt (5,5 x 8,5 Inch)
Private Const DMPAPER_EXECUTIVE               As Long = 7  ' Executive Blatt (7.25 x 10,5 Inch)
Private Const DMPAPER_A3                      As Long = 8  ' DIN A3 Blatt
Private Const DMPAPER_A4                      As Long = 9  ' DIN A4 Blatt
Private Const DMPAPER_A4SMALL                 As Long = 10 ' DIN A4 klein (210 x 297 Millimeter) Blatt
Private Const DMPAPER_A5                      As Long = 11 ' DIN A5 Blatt
Private Const DMPAPER_B4                      As Long = 12 ' B4 Blatt
Private Const DMPAPER_B5                      As Long = 13 ' B5 Blatt
Private Const DMPAPER_FOLIO                   As Long = 14 ' Folin 8,5 + 13 Inch
Private Const DMPAPER_QUARTO                  As Long = 15 ' Quarto Blatt (215 x 275 Millimeter)
Private Const DMPAPER_10X14                   As Long = 16 ' 10 x 14 Inch
Private Const DMPAPER_11X17                   As Long = 17 ' 11 x 17 Inch
Private Const DMPAPER_NOTE                    As Long = 18 ' Note Blatt (8,5 x 11 Inch)
Private Const DMPAPER_ENV_9                   As Long = 19 ' Briefumschlag 9 (3,875 x 8,875 Inch)
Private Const DMPAPER_ENV_10                  As Long = 20 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_11                  As Long = 21 ' Briefumschlag 11 (4,5 x 10,375 Inch)
Private Const DMPAPER_ENV_12                  As Long = 22 ' Briefumschlag 12 (4,75 x 11 Inch)
Private Const DMPAPER_ENV_14                  As Long = 23 ' Briefumschlag 14 (5 x 11,5 Inch)
Private Const DMPAPER_CSHEET                  As Long = 24 ' C Blatt (17 x 22 Inch)
Private Const DMPAPER_DSHEET                  As Long = 25 ' D Blatt (22x34 Inch)
Private Const DMPAPER_ESHEET                  As Long = 26 ' E Blatt (34 x 44 Inch)
Private Const DMPAPER_ENV_DL                  As Long = 27 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_C5                  As Long = 28 ' Briefumschlag C5 (162 x 229 Millimeter)
Private Const DMPAPER_ENV_C3                  As Long = 29 ' Briefumschlag C3 (324 x 458 Millimeter)
Private Const DMPAPER_ENV_C4                  As Long = 30 ' Briefumschlag C4 (229 x 324 Millimeter)
Private Const DMPAPER_ENV_C6                  As Long = 31 ' Briefumschlag C6 (114 x 162 Millimeter)
Private Const DMPAPER_ENV_C65                 As Long = 32 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_B4                  As Long = 33 ' Briefumschlag B4 (250 x 353 Millimeter)
Private Const DMPAPER_ENV_B5                  As Long = 34 ' Briefumschlag B5 (176 x 250 Millimeter)
Private Const DMPAPER_ENV_B6                  As Long = 35 ' Briefumschlag B6 (176 x 125 Millimeter)
Private Const DMPAPER_ENV_ITALY               As Long = 36 ' Italienischer Briefumschlag (110 x 230 Millimeter)
Private Const DMPAPER_ENV_MONARCH             As Long = 37 ' Monarchischer Briefumschlag (3,875 x 7,5 ' Inch)
Private Const DMPAPER_ENV_PERSONAL            As Long = 38 ' Persönlicher Briefumschlag (3,625 x 6,5 ' Inch)
Private Const DMPAPER_FANFOLD_US              As Long = 39 ' US Standard Fanfold (14,875 + 11 Inch)
Private Const DMPAPER_FANFOLD_STD_GERMAN      As Long = 40 ' Deutscher Standard Fanfold (8,5 x 12 ' Inch)
Private Const DMPAPER_FANFOLD_LGL_GERMAN      As Long = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_LAST                    As Long = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_ISO_B4                  As Long = 42 ' ISO B4 Blatt
Private Const DMPAPER_JAPANESE_POSTCARD       As Long = 43 ' Japanische Postkarte
Private Const DMPAPER_9X11                    As Long = 44 ' 9 x 11 Inch
Private Const DMPAPER_10X11                   As Long = 45 ' 10 x 11 Inch
Private Const DMPAPER_15X11                   As Long = 46 ' 15 x 11 Inch
Private Const DMPAPER_ENV_INVITE              As Long = 47 ' Einladungs Briefumschlag

Private Const DMPAPER_LETTER_EXTRA            As Long = 50 ' Blatt Extra
Private Const DMPAPER_LEGAL_EXTRA             As Long = 51 ' Rechtlich Extra
Private Const DMPAPER_TABLOID_EXTRA           As Long = 52 ' Tabloid Extra Blatt
Private Const DMPAPER_A4_EXTRA                As Long = 53 ' DIN A4 Extra Blatt
Private Const DMPAPER_LETTER_TRANSVERSE       As Long = 54 ' Blatt querlaufend
Private Const DMPAPER_A4_TRANSVERSE           As Long = 55 ' DIN A4 querlaufend
Private Const DMPAPER_LETTER_EXTRA_TRANSVERSE As Long = 56  ' Blatt Extra querlaufend
Private Const DMPAPER_A_PLUS                  As Long = 57 ' A Plus Blatt
Private Const DMPAPER_B_PLUS                  As Long = 58 ' B Plus Blatt
Private Const DMPAPER_LETTER_PLUS             As Long = 59 ' Blatt Plus
Private Const DMPAPER_A4_PLUS                 As Long = 60 ' DIN A4 Plus Blatt
Private Const DMPAPER_A5_TRANSVERSE           As Long = 61 ' DIN A5 querlaufend
Private Const DMPAPER_B5_TRANSVERSE           As Long = 62 ' B5 Blatt querlaufend
Private Const DMPAPER_A3_EXTRA                As Long = 63 ' DIN A3 Extra Blatt
Private Const DMPAPER_A5_EXTRA                As Long = 64 ' DIN A5 Extra Blatt
Private Const DMPAPER_B5_EXTRA                As Long = 65 ' B5 Extra Blatt
Private Const DMPAPER_A2                      As Long = 66 ' DIN A2 Blatt
Private Const DMPAPER_A3_TRANSVERSE           As Long = 67 ' DIN A3 Blatt querlaufend
Private Const DMPAPER_A3_EXTRA_TRANSVERSE     As Long = 68 ' DIN A3 Extra Blatt querlaufend
Private Const DMPAPER_USER                    As Long = 256 ' benutzerdefinierte Größe
 
' DEVMODE dmDefaultSource-Konstanten
Private Const DMBIN_ONLYONE       As Long = 1 ' nur ein Blatt
Private Const DMBIN_UPPER         As Long = 1  ' Behälter Oben
Private Const DMBIN_LOWER         As Long = 2 ' Behälter Unten
Private Const DMBIN_MIDDLE        As Long = 3 ' Behälter Mitte
Private Const DMBIN_MANUAL        As Long = 4 ' Behälter mit manueller Füllung
Private Const DMBIN_ENVELOPE      As Long = 5  ' Briefumschlag Behälter
Private Const DMBIN_ENVMANUAL     As Long = 6  ' Briefumschlag Behälter mit Manueller Füllung
Private Const DMBIN_AUTO          As Long = 7 ' automatisches füllen
Private Const DMBIN_TRACTOR       As Long = 8 ' Blatt Einzug
Private Const DMBIN_SMALLFMT      As Long = 9  ' kleines Format laden
Private Const DMBIN_LARGEFMT      As Long = 10 ' großes Format laden
Private Const DMBIN_LARGECAPACITY As Long = 11  ' große Kapazität
Private Const DMBIN_CASSETTE      As Long = 14 ' Blatt Kassette
Private Const DMBIN_FORMSOURCE    As Long = 15  ' Form Papier Quelle

' DEVMODE dmPrintQuality-Konstanten
Private Const DMRES_HIGH           As Long = -4  ' hoch
Private Const DMRES_MEDIUM         As Long = -3 ' mittel
Private Const DMRES_LOW            As Long = -2 '  niedrig
Private Const DMRES_DRAFT          As Long = -1 ' Entwurf

Public Enum PrinterResolutionKind
    High = -4    '/// <summary>Hohe Auflösung.</summary>
    Medium '= -3 '/// <summary>Mittlere Auflösung.</summary>
    Low    '= -2 '/// <summary>Niedrige Auflösung.</summary>
    Draft  '= -1 '/// <summary>Auflösung in Entwurfsqualität.</summary>
    Custom '= 0  '/// <summary>Benutzerdefinierte Auflösung.</summary>
End Enum

' DEVMODE dmColor-Konstanten
Private Const DMCOLOR_MONOCHROME   As Long = 1  ' Gerät unterstützt keine Schwarzweißausgabe
Private Const DMCOLOR_COLOR        As Long = 2 ' Gerät unterstützt keine Farbausgabe

' DEVMODE dmDuplex-Konstanten
Private Const DMDUP_SIMPLEX        As Long = 1 ' einseitiges Drucken
Private Const DMDUP_VERTICAL       As Long = 2  ' beidseitig Drucken bei vertikalem Seitenwechsel
Private Const DMDUP_HORIZONTAL     As Long = 3 ' beidseitig Drucken bei horizontalem Seitenwechsel

' DEVMODE dmTTOption-Konstanten
Private Const DMTT_BITMAP          As Long = 1 ' druckt TrueType-Fonts als Grafiken
Private Const DMTT_DOWNLOAD        As Long = 2 ' druckt Downloaded TrueType-Fonts als Soft-Fonts
Private Const DMTT_SUBDEV          As Long = 4 ' druckt Geräte-Fonts statt TrueType-Fonts

' DEVMODE dmCollate-Konstanten
Private Const DMCOLLATE_FALSE      As Long = 0  ' Drucker kann keine Seiten sortieren
Private Const DMCOLLATE_TRUE       As Long = 1  ' Drucker kann Seiten sortieren

' DEVMODE dmDisplayFlags-Konstanten
Private Const DM_GRAYSCALE         As Long = 1  ' Gerät unterstützt keine Farben, Grautöne werden unterstützt
Private Const DM_INTERLACED        As Long = 2 ' Gerät unterstützt Farben

' DEVMODE dmICMMethod-Konstanten
Private Const DMICMMETHOD_NONE     As Long = 1 ' ICM ist abgeschaltet
Private Const DMICMMETHOD_SYSTEM   As Long = 2  ' ICM wird von Windows gesteuert
Private Const DMICMMETHOD_DRIVER   As Long = 3 ' ICM wird vom Treiber gesteuert
Private Const DMICMMETHOD_DEVICE   As Long = 4  ' ICM wird vom Gerät gesteuert

' DEVMODE dmICMIntent-Konstanten
Private Const DMICM_SATURATE       As Long = 1 ' Sättigung wird optimiert
Private Const DMICM_CONTRAST       As Long = 2 ' Kontrast wird optimiert
Private Const DMICM_COLORMETRIC    As Long = 3 ' exakte Farben

' DEVMODE dmMediaType-Konstanten
Private Const DMMEDIA_STANDARD     As Long = 1  ' Standardpapier
Private Const DMMEDIA_GLOSSY       As Long = 2 ' Fotopapier
Private Const DMMEDIA_TRANSPARECNY As Long = 3  ' Folie

' DEVMODE dmDitherType-Konstanten
Private Const DMDITHER_NONE        As Long = 1 ' kein Dithering
Private Const DMDITHER_COARSE      As Long = 2 ' Dithering mit grobem Brush
Private Const DMDITHER_FINE        As Long = 3 ' Dithering mit feinem Brush
Private Const DMDITHER_LINEART     As Long = 4 ' Linien mit schwarz, weiß und grau
Private Const DMDITHER_GRAYSCALE   As Long = 5  ' Grauskala

'--------------------------------------------------------------------------------------------------------

' GlobalAlloc wFlags-Konstanten
Private Const GHND                 As Long = &H40 ' Kombination von GMEM_MOVEABLE mit GMEM_ZEROINIT
Private Const GMEM_DDESHARE        As Long = &H2000 ' optimiert den Speicher mit der DDE
Private Const GMEM_DISCARDABLE     As Long = &H100  ' Speicher kann überschrieben werden, kann
' nicht mit GMEM_FIXED kombiniert werden
Private Const GMEM_FIXED           As Long = &H0 ' fixiert den Speicher und die Funktion gibt einen
' Pointer zurück
Private Const GMEM_MOVEABLE        As Long = &H2 ' Der Speicher ist beweglich
Private Const GMEM_NOCOMPACT       As Long = &H10 ' Speicher wird nicht minimiert oder andere
' Blöcke für ihn überschrieben
Private Const GMEM_NODISCARD       As Long = &H20 ' überschreibt keine überschreibbaren Blöcke um
' den Speicher zu reservieren
Private Const GMEM_ZEROINIT        As Long = &H40  ' initialisiert den Inhalt des Speicherblocks bei 0
Private Const GPTR                 As Long = &H42 ' kombiniert GMEM_FIXED und GEMEM_ZEROINIT

'Enum PrinterOrientationConstants
'    cdlPortrait = 1
'    cdlLandscape = 2

'Oh wowo if you deal with PrintDialog you come across a whole new universe of printer settings and print settings, paper size and so many more other stuff


#If VBA7 Then
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long

    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)
    
    #If IsUnicode Then
        Private Declare PtrSafe Function PrintDlg Lib "comdlg32" Alias "PrintDlgW" (pPrintDlg As TPrintDialog) As Long
        'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
        Private Declare PtrSafe Function PrintDlgEx Lib "comdlg32" Alias "PrintDlgExW" (ByRef TLPPRINTDLGEX As TPrintDialogEx) As Long
    #Else
        Private Declare PtrSafe Function PrintDlg Lib "comdlg32" Alias "PrintDlgA" (pPrintDlg As TPrintDialog) As Long
        'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
        Private Declare PtrSafe Function PrintDlgEx Lib "comdlg32" Alias "PrintDlgExA" (ByRef TLPPRINTDLGEX As TPrintDialogEx) As Long
    #End If
    'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
    Private Declare PtrSafe Function PrintDlgExW Lib "comdlg32" (ByRef TLPPRINTDLGEXW As TPrintDialogEx) As Long
    Private Declare PtrSafe Function PrinterMessageBoxW Lib "winspool" (ByVal hPrinter As LongPtr, ByVal error As Long, ByVal hwnd As LongPtr, ByVal pText As LongPtr, ByVal pCaption As LongPtr, ByVal dwType As Long) As Long
    Private Declare PtrSafe Function PrinterProperties Lib "winspool" (ByVal hwnd As LongPtr, ByVal hPrinter As LongPtr) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLen As Long)
    
    #If IsUnicode Then
        Private Declare Function PrintDlg Lib "comdlg32" Alias "PrintDlgW" (pPrintDlg As TPrintDialog) As Long
        'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
        Private Declare Function PrintDlgEx Lib "comdlg32" Alias "PrintDlgExW" (ByRef TLPPRINTDLGEXW As TPrintDialogEx) As Long
    #Else
        Private Declare Function PrintDlg Lib "comdlg32" Alias "PrintDlgA" (pPrintDlg As TPrintDialog) As Long
        'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
        Private Declare Function PrintDlgEx Lib "comdlg32" Alias "PrintDlgExA" (ByRef TLPPRINTDLGEXW As TPrintDialogEx) As Long
    #End If
    'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646940(v=vs.85)
    Private Declare Function PrintDlgExW Lib "comdlg32" (ByRef TLPPRINTDLGEXW As TPrintDialogEx) As Long
    Private Declare Function PrinterMessageBoxW Lib "winspool" (ByVal hPrinter As LongPtr, ByVal error As Long, ByVal hwnd As LongPtr, ByVal pText As LongPtr, ByVal pCaption As LongPtr, ByVal dwType As Long) As Long
    Private Declare Function PrinterProperties Lib "winspool" (ByVal hwnd As LongPtr, ByVal hPrinter As LongPtr) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
#End If

'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646942(v=vs.85)
'https://docs.microsoft.com/de-de/windows/win32/api/commdlg/ns-commdlg-printdlgexw

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

Private Type TPrintDialog
    lStructSize   As Long
    hwndOwner     As LongPtr
    hDevMode      As LongPtr
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

Private Type TPrintDialogEx
    lStructSize      As Long
    hwndOwner        As LongPtr '  HWND
    hDevMode         As LongPtr '  HGLOBAL
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
    pszHeaderTitle    As LongPtr ' LPCWSTR
    pszHeaderSubTitle As LongPtr ' LPCWSTR
    hActCtx           As LongPtr ' HANDLE
'  union {
'    HBITMAP hbmHeader;
'    LPCWSTR pszbmHeader;
'  } DUMMYUNIONNAME3;
    union_hbmHeader_pszbmHeader As LongPtr
End Type

Private Type TDeviceNames
    wDriverOffset    As Integer      '   2
    wDeviceOffset    As Integer      '   2
    wOutputOffset    As Integer      '   2
    wDefault         As Integer      '   2
    'extra(0 To 99)   As Byte 'String * 100
    extra            As String * 100 ' 100
    'extra(0 To 100 - 1) As Byte '* 100
End Type                        ' Sum: 108

Private Type TDeviceMode
    dmDeviceName     As String * 32 ' 32
    dmSpecVersion    As Integer     '  2
    dmDriverVersion  As Integer     '  2
    dmSize           As Integer     '  2
    dmDriverExtra    As Integer     '  2
    dmFields         As Long        '  4
    dmOrientation    As Integer     '  2
    dmPaperSize      As Integer     '  2
    dmPaperLength    As Integer     '  2
    dmPaperWidth     As Integer     '  2
    dmScale          As Integer     '  2
    dmCopies         As Integer     '  2
    dmDefaultSource  As Integer     '  2
    dmPrintQuality   As Integer     '  2
    dmColor          As Integer     '  2
    dmDuplex         As Integer     '  2
    dmYResolution    As Integer     '  2
    dmTTOption       As Integer     '  2
    dmCollate        As Integer     '  2
    dmFormName       As String * 32 ' 32
    dmUnusedPadding  As Integer     '  2
    dmBitsPerPixel   As Integer     '  2
    dmPelsWidth      As Long        '  4
    dmPelsHeight     As Long        '  4
    dmDisplayFlags   As Long        '  4
    dmDisplayFrequency As Long      '  4
    dmICMMethod      As Long        '  4
    dmICMIntent      As Long        '  4
    dmMediaType      As Long        '  4
    dmDitherType     As Long        '  4
    dmReserved1      As Long        '  4
    dmReserved2      As Long        '  4
    ' der folgende Abschnitt ist nur für Windows 2000 relevant
    dmPanningWidth   As Long        '  4
    dmPanningHeight  As Long        '  4
End Type                      ' Sum: 154

Private m_PrintDialog   As TPrintDialog
Private m_UseEXDialog   As Boolean
Private m_PrintDialogEx As TPrintDialogEx
Private m_DeviceMode    As TDeviceMode
Private m_DeviceNames   As TDeviceNames

Private m_DriverName  As String
Private m_DeviceName  As String
Private m_OutputName  As String
Private m_DefaultName As String

Private mFlags        As Long

Private Sub Class_Initialize()
    m_PrintDialog.lStructSize = LenB(m_PrintDialog) '66' 110 '120
    m_PrintDialogEx.lStructSize = LenB(m_PrintDialogEx) '120 '136
    'PD_DISABLEPRINTTOFILE Or
    mFlags = PD_COLLATE Or PD_NOSELECTION 'PD_RETURNDC Or PD_ALLPAGES Or PD_PAGENUMS Or PD_USEDEVMODECOPIESANDCOLLATE
End Sub

'Properties of Windows.Forms.PrintDialog
Public Property Get AllowCurrentPage() As Boolean
    'only togehter with UseExDialog = True
    'AllowCurrentPage = (mFlags And PD_ALLPAGES) = 0
End Property
Public Property Let AllowCurrentPage(ByVal Value As Boolean)
    'mFlags and
End Property

Public Property Get AllowPrintToFile() As Boolean
    'AllowPrintToFile = m_PrintDialog.flags And PD_DISABLEPRINTTOFILE = 0
    AllowPrintToFile = (mFlags And PD_DISABLEPRINTTOFILE) <> PD_DISABLEPRINTTOFILE
End Property
Public Property Let AllowPrintToFile(ByVal Value As Boolean)
'    AllowPrintToFile = m_PrintDialog.flags Or PD_DISABLEPRINTTOFILE
    mFlags = mFlags Or PD_DISABLEPRINTTOFILE
    If Value Then mFlags = mFlags Xor PD_DISABLEPRINTTOFILE
End Property

'Public Property Get CheckFileExists() As Boolean              'OFN_FILEMUSTEXIST
'    CheckFileExists = mFlags And OFN_FILEMUSTEXIST
'End Property
'Public Property Let CheckFileExists(ByVal Value As Boolean)   'OFN_FILEMUSTEXIST
'    mFlags = mFlags Or OFN_FILEMUSTEXIST
'    If Not Value Then mFlags = mFlags Xor OFN_FILEMUSTEXIST
'End Property

Public Property Get AllowSelection() As Boolean
    AllowSelection = (mFlags And PD_NOSELECTION) <> PD_NOSELECTION
End Property
Public Property Let AllowSelection(ByVal Value As Boolean)
    mFlags = mFlags Or PD_NOSELECTION
    If Value Then mFlags = mFlags Xor PD_NOSELECTION
End Property

Public Property Get AllowSomePages() As Boolean
    AllowSomePages = (mFlags And PD_NOPAGENUMS) <> PD_NOPAGENUMS
End Property
Public Property Let AllowSomePages(ByVal Value As Boolean)
    mFlags = mFlags Or PD_NOPAGENUMS
    If Not Value Then mFlags = mFlags Xor PD_NOPAGENUMS
End Property

Public Property Get UseEXDialog() As Boolean
    UseEXDialog = m_UseEXDialog
End Property
Public Property Let UseEXDialog(ByVal Value As Boolean)
    m_UseEXDialog = Value
End Property

'Public Property Get ShowNetwork() As Boolean
'    ShowNetwork = Not CBool((mFlags And PD_NONETWORKBUTTON))
'End Property
'Public Property Let ShowNetwork(ByVal Value As Boolean)
'    mFlags = mFlags Or PD_NONETWORKBUTTON
'    If Value Then mFlags = mFlags Xor PD_NONETWORKBUTTON
'End Property

Public Property Get ShowNetwork() As Boolean
    ShowNetwork = (m_PrintDialog.flags And PD_NONETWORKBUTTON) = 0
End Property
Public Property Let ShowNetwork(ByVal Value As Boolean)
    ShowNetwork = (m_PrintDialog.flags And PD_NONETWORKBUTTON) = 0
End Property

'Public Property Get Container() As IContainer
'    '
'End Property

' v ############################## v '        Document       ' v ############################## v '
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
' ^ ############################## ^ '        Document       ' ^ ############################## ^ '

' v ############################## v '    PrinterSettings    ' v ############################## v '

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
    PrinterSettings_Copies = m_PrintDialog.nCopies
End Property
Public Property Let PrinterSettings_Copies(ByVal Value As Integer)
    m_PrintDialog.nCopies = Value
End Property

'Public Property Get PrinterSettings_DefaultPageSettings_Bounds() As Rectangle
'    '
'End Property
'Public Property Get PrinterSettings_DefaultPageSettings_Clone() As Object
    '
'End Property
Public Property Get PrinterSettings_DefaultPageSettings_Color() As Boolean
    '
End Property
'Public Property Get PrinterSettings_DefaultPageSettings_CopyToHdevmode(hdevmode As IntPtr)
'    '
'End Property
'Public Property Get PrinterSettings_Duplex() As Duplex 'Enum
'    '
'End Property

' v ############################## v '    PrinterSettings.Duplex    ' v ############################## v '
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
' ^ ############################## ^ '    PrinterSettings.Duplex    ' ^ ############################## ^ '

' v ############################## v '    PrinterSettings    ' v ############################## v '
Public Property Get PrinterSettings_FromPage() As Integer
    PrinterSettings_FromPage = m_PrintDialog.nFromPage
End Property
Public Property Let PrinterSettings_FromPage(ByVal Value As Integer)
    m_PrintDialog.nFromPage = Value
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

Public Property Get PrinterSettings_MaximumPage() As Integer
    PrinterSettings_MaximumPage = m_PrintDialog.nMaxPage
End Property
Public Property Let PrinterSettings_MaximumPage(ByVal Value As Integer)
    m_PrintDialog.nMaxPage = Value
End Property

Public Property Get PrinterSettings_MinimumPage() As Integer
    PrinterSettings_MaximumPage = m_PrintDialog.nMinPage
End Property
Public Property Let PrinterSettings_MinimumPage(ByVal Value As Integer)
    m_PrintDialog.nMinPage = Value
End Property

'Public Property Get PrinterSettings_PaperSizes() As PaperSizeCollection
'    '
'End Property
'Public Property Get PrinterSettings_PaperSources() As PaperSourceCollection
'    '
'End Property

Public Property Get PrinterSettings_PrinterName() As String
    PrinterSettings_PrinterName = m_DeviceName
End Property
Public Property Get PrinterSettings_PrinterDriverName() As String
    PrinterSettings_PrinterDriverName = m_DriverName
End Property
Public Property Get PrinterSettings_PrinterOutputName() As String
    PrinterSettings_PrinterOutputName = m_OutputName
End Property
Public Property Get PrinterSettings_PrinterDefaultName() As String
    PrinterSettings_PrinterDefaultName = m_DefaultName
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

'Public Property Get CheckFileExists() As Boolean              'OFN_FILEMUSTEXIST
'    CheckFileExists = mFlags And OFN_FILEMUSTEXIST
'End Property
'Public Property Let CheckFileExists(ByVal Value As Boolean)   'OFN_FILEMUSTEXIST
'    mFlags = mFlags Or OFN_FILEMUSTEXIST
'    If Not Value Then mFlags = mFlags Xor OFN_FILEMUSTEXIST
'End Property


Public Property Get PrinterSettings_PrintToFile() As Boolean
    PrinterSettings_PrintToFile = (mFlags And PD_PRINTTOFILE) = PD_PRINTTOFILE
End Property

Public Property Let PrinterSettings_PrintToFile(ByVal Value As Boolean)
    mFlags = mFlags Or PD_PRINTTOFILE
    If Not Value Then mFlags = mFlags Xor PD_PRINTTOFILE
End Property


Public Property Get PrinterSettings_SupportsColor() As Boolean
    '
End Property

Public Property Get PrinterSettings_ToPage() As Integer
    PrinterSettings_ToPage = m_PrintDialog.nToPage
End Property
Public Property Let PrinterSettings_ToPage(ByVal Value As Integer)
    m_PrintDialog.nToPage = Value
End Property
' ^ ############################## ^ '    PrinterSettings    ' ^ ############################## ^ '

Public Property Get PrintToFile() As Boolean
    PrintToFile = (mFlags And PD_PRINTTOFILE) = PD_PRINTTOFILE
End Property

Public Sub Reset()
    '
End Sub

Private Function App_hInstance() As LongPtr
#If VBA7 Then
    App_hInstance = Excel.Application.hInstancePtr
#ElseIf VBA6 Then
    App_hInstance = Excel.Application.hInstance
#Else
    App_hInstance = App.hInstance
#End If
End Function

Private Function GethWnd(FOwner) As LongPtr
#If VBA7 Or VBA6 Then
    GethWnd = GetDesktopWindow
#Else
    GethWnd = FOwner.hwnd
#End If
End Function

Public Function ShowDialog(Optional FOwner As Variant) As VbMsgBoxResult
'Public Function ShowDialog() As VbMsgBoxResult

'https://microsoft.public.de.vb.narkive.com/wz5CVka1/druckereinstellung-mit-printdialog-oder-commondialog-showprinter
Try: On Error GoTo Catch

    ' Drucker festlegen, der standardmässig ausgewählt sein soll
    With m_DeviceMode
        .dmSize = LenB(m_DeviceMode) '154
        '.dmDeviceName = Printer.DeviceName
        .dmFields = DM_PRINTQUALITY
    End With
    
    ' Speicherblock erstellen und Dev_Mode hineinkopieren
    Dim hMemMode As LongPtr: hMemMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(m_DeviceMode)) 'Len(Dev_Mode)=154
    ' Speicherblock erstellen
    Dim pMemMode As LongPtr: pMemMode = GlobalLock(hMemMode) ' Speicherblock sperren und Pointer ermitteln
    RtlMoveMemory ByVal pMemMode, m_DeviceMode, LenB(m_DeviceMode)  ' Struktur in den Speicherblock kopieren
    GlobalUnlock hMemMode ' Speicherblock entsperren
    
    ' Drucker-Eigenschaften zur Überprüfung festlegen
    With m_DeviceNames
        '.wDriverOffset = Len(Printer.DriverName)
        '.wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        '.wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        Dim s As String ': s = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.Port & vbNullChar
        's = StrConv(s, vbNarrow)
        
        .extra = s
        'RtlMoveMemory VarPtr(Dev_Names) + 8, StrPtr(s), Len(s)
    End With
    
    ' Speicherblock erstellen und Dev_Names hineinkopieren
    Dim hMemName As LongPtr: hMemName = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(m_DeviceNames))
    ' Speicherblock erstellen
    Dim pMemName As LongPtr: pMemName = GlobalLock(hMemName) ' Speicherblock sperren und Pointer ermitteln
    RtlMoveMemory ByVal pMemName, m_DeviceNames, LenB(m_DeviceNames) ' Struktur in den Speicherblock kopieren
    GlobalUnlock hMemName ' Speicherblock entsperren

    Dim hr As Long
    If m_UseEXDialog Then
        With m_PrintDialogEx
            .flags = mFlags
        End With
        hr = PrintDlgEx(m_PrintDialogEx)
    Else
        With m_PrintDialog
            '.hwndOwner = GethWnd(FOwner)
            .hInstance = App_hInstance
            
            .hDevMode = hMemMode
            .hDevNames = hMemName
            '.flags = mFlags
            .flags = PD_DISABLEPRINTTOFILE Or PD_RETURNDC Or PD_ALLPAGES Or PD_PAGENUMS Or PD_USEDEVMODECOPIESANDCOLLATE
            
            '.flags = PD_USEDEVMODECOPIESANDCOLLATE Or PD_RETURNDC
            '.nCopies = 1
            '.nFromPage = 1 '0 '&HFFFF
            '.nToPage = 100 '&HFFFF
            '.nMinPage = 1
            '.nMaxPage = 100 '&HFFFF
        End With
        hr = PrintDlg(m_PrintDialog)
        mFlags = m_PrintDialog.flags
    End If
    ShowDialog = IIf(hr, vbOK, vbCancel)
    If ShowDialog = vbOK Then
        ' Dev_Mode wieder aus dem Speicherblock in die Struktur kopieren und
        ' Speicherblock zerstören
        pMemMode = GlobalLock(m_PrintDialog.hDevMode) ' Speicherblock sperren
        RtlMoveMemory m_DeviceMode, ByVal pMemMode, LenB(m_DeviceMode) ' Daten aus dem
        ' Speicherblock in die Struktur kopieren
        GlobalUnlock m_PrintDialog.hDevMode ' Speicherblock entsperren
        'GlobalFree PD.hDevMode ' Speicherblock zerstören
        
        GetNames m_PrintDialog.hDevNames
        ' Dev_Names wieder aus dem Speicherblock in die Struktur kopieren und
        ' Speicherblock zerstören
        pMemName = GlobalLock(m_PrintDialog.hDevNames)  ' Speicherblock sperren
        RtlMoveMemory m_DeviceNames, ByVal pMemName, LenB(m_DeviceNames) ' Daten aus dem
        ' Speicherblock in die Struktur kopieren
        GlobalUnlock m_PrintDialog.hDevNames ' Speicherblock entsperren
        'GlobalFree PD.hDevNames ' Speicherblock zerstören
    End If
    
    GoTo Finally
Catch:
    MsgBox "Error calling PrintDialog"
Finally:
    GlobalFree m_PrintDialog.hDevMode  ' Speicher freigeben
    GlobalFree m_PrintDialog.hDevNames ' Speicher freigeben
    ' Device-Kontext freigeben
    DeleteObject m_PrintDialog.hDC
End Function


'Beispiel: Druckernamen der gewählten Einstellungen ermitteln
' (Geht so auch bei hDevMode->DevMode-Struktur)
Private Sub GetNames(ByVal hDevNames As LongPtr) 'As String
    
    'Struktur anhand des Handles füllen (Kopieren)
    Dim LenDevNames As Long: LenDevNames = GlobalSize(hDevNames)
    Dim pDevNames As LongPtr:  pDevNames = GlobalLock(hDevNames)
    Dim DN As TDeviceNames 'DEVNAMES
    RtlMoveMemory DN, ByVal pDevNames, LenDevNames
    GlobalUnlock hDevNames
    
    'Struktur auswerten
    With DN
        Dim ex As String: ex = StrConv(.extra, vbFromUnicode)
        'Dim ex As String: ex = .extra
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


Public Property Get ShowHelp() As Boolean
    ShowHelp = (mFlags And PD_SHOWHELP) = PD_SHOWHELP
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    mFlags = mFlags Or PD_SHOWHELP
    If Not Value Then mFlags = mFlags Xor PD_SHOWHELP
End Property

'Public Property Get ShowNetwork() As Boolean
'    ShowNetwork = Not CBool((mFlags And PD_NONETWORKBUTTON))
'End Property
'Public Property Let ShowNetwork(ByVal Value As Boolean)
'    mFlags = mFlags Or PD_NONETWORKBUTTON
'    If Value Then mFlags = mFlags Xor PD_NONETWORKBUTTON
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


