Attribute VB_Name = "MPrinterPaper"
Option Explicit

Public Enum PaperKind
    Custom = 256                            ' Das Papierformat wird vom Benutzer festgelegt.
    Letter = 1                              ' Letter paper (8.5 in.by 11 in.).
    LetterSmall = 2                         ' Letter small paper (8.5 in.by 11 in.).
    Tabloid = 3                             ' Tabloid paper (11 in.by 17 in.).
    Ledger = 4                              ' Ledger paper (17 in.by 11 in.).
    Legal = 5                               ' Legal paper (8.5 in.by 14 in.).
    Statement = 6                           ' Statement paper (5.5 in.by 8.5 in.).
    Executive = 7                           ' Executive paper (7.25 in.by 10.5 in.).
    DIN_A3 = 8                              ' A3 paper (297 mm by 420 mm).
    DIN_A4 = 9                              ' A4 (210 x 297 mm).
    DIN_A4Small = 10                        ' A4 klein (210 x 297 mm).
    DIN_A5 = 11                             ' A5 (148 x 210 mm).
    DIN_B4 = 12                             ' B4 (250 x 353 mm).
    DIN_B5 = 13                             ' B5 (176 x 250 mm).
    Folio = 14                              ' Folio paper (8.5 in.by 13 in.).
    Quarto = 15                             ' Quarto (215 x 275 mm).
    Standard10x14 = 16                      ' Standard paper (10 in.by 14 in.).
    Standard11x17 = 17                      ' Standard paper (11 in.by 17 in.).
    Note = 18                               ' Note paper (8.5 in.by 11 in.).
    Number9Envelope = 19                    ' #9 envelope (3.875 in.by 8.875 in.).
    Number10Envelope = 20                   ' #10 envelope (4.125 in.by 9.5 in.).
    Number11Envelope = 21                   ' #11 envelope (4.5 in.by 10.375 in.).
    Number12Envelope = 22                   ' #12 envelope (4.75 in.by 11 in.).
    Number14Envelope = 23                   ' #14 envelope (5 in.by 11.5 in.).
    CSheet = 24                             ' C paper (17 in.by 22 in.).
    DSheet = 25                             ' D paper (22 in.by 34 in.).
    ESheet = 26                             ' E paper (34 in.by 44 in.).
    DLEnvelope = 27                         ' Umschlag DL (110 x 220 mm).
    DIN_C5Envelope = 28                     ' Umschlag C5 (162 x 229 mm).
    DIN_C3Envelope = 29                     ' Umschlag C3 (324 x 458 mm).
    DIN_C4Envelope = 30                     ' Umschlag C4 (229 x 324 mm).
    DIN_C6Envelope = 31                     ' C6 envelope (114 mm by 162 mm).
    DIN_C65Envelope = 32                    ' C65 envelope (114 mm by 229 mm).
    DIN_B4Envelope = 33                     ' B4 (250 x 353 mm).
    DIN_B5Envelope = 34                     ' Umschlag B5 (176 x 250 mm).
    DIN_B6Envelope = 35                     ' Umschlag B6 (176 x 125 mm).
    ItalyEnvelope = 36                      ' Umschlag Italien (110 x 230 mm).
    MonarchEnvelope = 37                    ' Monarch envelope (3.875 in.by 7.5 in.).
    PersonalEnvelope = 38                   ' 6 3/4 envelope (3.625 in.by 6.5 in.).
    USStandardFanfold = 39                  ' US standard fanfold (14.875 in.by 11 in.).
    GermanStandardFanfold = 40              ' German standard fanfold (8.5 in.by 12 in.).
    GermanLegalFanfold = 41                 ' German legal fanfold (8.5 in.by 13 in.).
    DIN_IsoB4 = 42                          ' B4 (ISO) (250 x 353 mm).
    JapanesePostcard = 43                   ' Japanische Postkarte (100 x 148 mm).
    Standard9x11 = 44                       ' Standard paper (9 in.by 11 in.).
    Standard10x11 = 45                      ' Standard paper (10 in.by 11 in.).
    Standard15x11 = 46                      ' Standard paper (15 in.by 11 in.).
    InviteEnvelope = 47                     ' Einladungsumschlag (220 x 220 mm).
    '? = 48
    '? = 49
    LetterExtra = 50                        ' Letter extra paper (9.275 in.by 12 in.).Dieser Wert ist PostScript-Treiber-spezifisch und wird ausschließlich von Linotronic-Druckern zur Senkung des Papierverbrauchs verwendet.
    LegalExtra = 51                         ' Legal extra paper (9.275 in.by 15 in.).Dieser Wert ist PostScript-Treiber-spezifisch und wird ausschließlich von Linotronic-Druckern zur Senkung des Papierverbrauchs verwendet.
    TabloidExtra = 52                       ' Tabloid extra paper (11.69 in.by 18 in.).Dieser Wert ist PostScript-Treiber-spezifisch und wird ausschließlich von Linotronic-Druckern zur Senkung des Papierverbrauchs verwendet.
    DIN_A4Extra = 53                            ' A4 Extra (236 x 322 mm).Dieser Wert ist PostScript-Treiber-spezifisch und wird ausschließlich von Linotronic-Druckern zur Senkung des Papierverbrauchs verwendet.
    LetterTransverse = 54                   ' Letter transverse paper (8.275 in.by 11 in.).
    DIN_A4Transverse = 55                       ' A4 transverse paper (210 mm by 297 mm).
    LetterExtraTransverse = 56              ' Letter extra transverse paper (9.275 in.by 12 in.).
    APlus = 57                              ' SuperA/SuperA/A4 (227 x 356 mm).
    BPlus = 58                              ' SuperB/SuperB/A3 (305 x 487 mm).
    LetterPlus = 59                         ' Letter plus paper (8.5 in.by 12.69 in.).
    DIN_A4Plus = 60                         ' A4 Plus (210 x 330 mm).
    DIN_A5Transverse = 61                   ' A5 gedreht (148 x 210 mm).
    DIN_B5Transverse = 62                   ' B5 (JIS) gedreht (182 x 257 mm).
    DIN_A3Extra = 63                        ' A3 extra paper (322 mm by 445 mm).
    DIN_A5Extra = 64                        ' A5 Extra (174 x 235 mm).
    DIN_B5Extra = 65                        ' B5 (ISO) Extra (201 x 276 mm).
    DIN_A2 = 66                             ' A2 paper (420 mm by 594 mm).
    DIN_A3Transverse = 67                   ' A3 transverse paper (297 mm by 420 mm).
    DIN_A3ExtraTransverse = 68                  ' A3 Extra quer (322 x 445 mm).
    JapaneseDoublePostcard = 69             ' Japanische Doppelpostkarte (200 x 148 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_A6 = 70                             ' A6 (105 x 148 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeKakuNumber2 = 71        ' Japanischer Umschlag Kaku #2.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeKakuNumber3 = 72        ' Japanischer Umschlag Kaku #3.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeChouNumber3 = 73        ' Japanischer Umschlag Chou #3.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeChouNumber4 = 74        ' Japanischer Umschlag Chou #4.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    LetterRotated = 75                      ' Letter rotated paper (11 in.by 8.5 in.).
    DIN_A3Rotated = 76                      ' A3 gedreht (420 x 297 mm).
    DIN_A4Rotated = 77                      ' A4 gedreht (297 x 210 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_A5Rotated = 78                      ' A5 rotated paper (210 mm by 148 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_B4JisRotated = 79                   ' JIS B4 rotated paper (364 mm by 257 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_B5JisRotated = 80                   ' B5 (JIS) gedreht (257 x 182 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapanesePostcardRotated = 81            ' Japanische Postkarte gedreht (148 x 100 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseDoublePostcardRotated = 82      ' Japanische Doppelpostkarte gedreht (148 x 200 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_A6Rotated = 83                      ' A6 gedreht (148 x 105 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeKakuNumber2Rotated = 84 ' Japanischer Umschlag Kaku #2 gedreht.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeKakuNumber3Rotated = 85 ' Japanischer Umschlag Kaku #3 gedreht.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeChouNumber3Rotated = 86 ' Japanischer Umschlag Chou #3 gedreht.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeChouNumber4Rotated = 87 ' Japanischer Umschlag Chou #4 gedreht.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_B6Jis = 88                          ' B6 (JIS) (128 x 182 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    DIN_B6JisRotated = 89                   ' B6 (JIS) gedreht (182 x 128 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Standard12x11 = 90                      ' Standard paper (12 in.by 11 in.).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeYouNumber4 = 91         ' Japanischer Umschlag You #4.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    JapaneseEnvelopeYouNumber4Rotated = 92  ' Japanischer Umschlag You #4 gedreht.Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc16K = 93                             ' Volksrepublik China 16K (146 x 215 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc32K = 94                             ' Volksrepublik China 32K (97 x 151 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc32KBig = 95                          ' Volksrepublik China 32K groß (97 x 151 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber1 = 96                 ' Volksrepublik China #1 Umschlag (102 x 165 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber2 = 97                 ' Volksrepublik China #2 Umschlag (102 x 176 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber3 = 98                 ' Volksrepublik China #3 Umschlag (125 x 176 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber4 = 99                 ' Volksrepublik China #4 Umschlag (110 x 208 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber5 = 100                ' Volksrepublik China #5 Umschlag (110 x 220 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber6 = 101                ' Volksrepublik China #6 Umschlag (120 x 230 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber7 = 102                ' Volksrepublik China #7 Umschlag (160 x 230 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber8 = 103                ' Volksrepublik China #8 Umschlag (120 x 309 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber9 = 104                ' Volksrepublik China #9 Umschlag (229 x 324 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber10 = 105               ' Volksrepublik China #10 Umschlag (324 x 458 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc16KRotated = 106                     ' Volksrepublik China 16K gedreht (146 x 215 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc32KRotated = 107                     ' Volksrepublik China 32K gedreht (97 x 151 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    Prc32KBigRotated = 108                  ' Volksrepublik China 32K groß gedreht (97 x 151 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber1Rotated = 109         ' Volksrepublik China #1 Umschlag gedreht (165 x 102 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber2Rotated = 110         ' Volksrepublik China #2 Umschlag gedreht (176 x 102 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber3Rotated = 111         ' Volksrepublik China #3 Umschlag gedreht (176 x 125 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber4Rotated = 112         ' Volksrepublik China #4 Umschlag gedreht (208 x 110 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber5Rotated = 113         ' Volksrepublik China #5 Umschlag gedreht (220 x 110 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber6Rotated = 114         ' Volksrepublik China #6 Umschlag gedreht (230 x 120 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber7Rotated = 115         ' Volksrepublik China #7 Umschlag gedreht (230 x 160 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber8Rotated = 116         ' Volksrepublik China #8 Umschlag gedreht (309 x 120 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber9Rotated = 117         ' Volksrepublik China #9 Umschlag gedreht (324 x 229 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
    PrcEnvelopeNumber10Rotated = 118        ' Volksrepublik China #10 Umschlag gedreht (458 x 324 mm).Erfordert Windows 98, Windows NT 4.0 oder eine höhere Version.
End Enum

Public Type PaperSize
    Height    As Long
    Width     As Long
    Kind      As PaperKind
    RawKind   As Long
    PaperName As String
End Type

'defined in comdlg32.ocx
'Public Enum PaperSourceKind
'    Upper = 1          ' Das obere Fach eines Druckers (oder das Standardfach bei einem Drucker mit nur einem Fach).
'    Lower = 2          ' Das untere Fach eines Druckers.
'    Middle = 3         ' Das mittlere Fach eines Druckers.
'    Manual = 4         ' Manuell zugeführtes Papier.
'    Envelope = 5       ' Ein Briefumschlag.
'    ManualFeed = 6     ' Manuell zugeführter Briefumschlag.
'    AutomaticFeed = 7  ' Automatischer Papiereinzug.
'    TractorFeed = 8    ' Ein Traktoreinzug.
'    SmallFormat = 9    ' Kleinformatiges Papier.
'    LargeFormat = 10   ' Großformatiges Papier.
'    LargeCapacity = 11 ' Das Druckerfach mit großer Kapazität.
'    Cassette = 14      ' Eine Papierkassette.
'    FormSource = 15    ' Das Standardzufuhrfach des Druckers.
'    Custom = 257       ' Druckerspezifische Papierzufuhr.
'End Enum

'Enum PrinterConstants
'    cdlPDAllPages = 0         '         ' Gibt den Zustand des Optionsfeldes 'Alles' zurück oder legt ihn fest.
'    cdlPDSelection = 1        '         ' Gibt den Zustand des Optionsfeldes 'Auswahl' zurück oder legt ihn fest.
'    cdlPDPageNums = 2         '         ' Gibt den Zustand des Optionsfeldes 'Seiten' zurück oder legt ihn fest.
'    cdlPDNoSelection = 4      '         ' Sperrt das Optionsfeld 'Auswahl'.
'    cdlPDNoPageNums = 8       '         ' Gibt den Zustand des Optionsfeldes 'Seiten' zurück oder legt ihn fest.
'    cdlPDCollate = 16         ' (&H10)  ' Gibt den Zustand des Feldes 'Exemplare' zurück oder legt ihn fest.
'    cdlPDPrintToFile = 32     ' (&H20)  ' Gibt den Zustand des Kontrollkästchens 'Ausdruck in Datei' zurück oder legt ihn fest.
'    cdlPDPrintSetup = 64      ' (&H40)  ' Zeigt das Dialogfeld 'Drucker einrichten' anstelle des Dialogfelds 'Drucken' an.
'    cdlPDNoWarning = 128      ' (&H80)  ' Unterbindet eine Warnmeldung, wenn es keinen Standarddrucker gibt.
'    cdlPDReturnDC = 256       ' (&H100) ' Gibt einen Gerätekontext für die Druckerauswahl zurück.
'    cdlPDReturnIC = 512       ' (&H200) ' Gibt einen Informationskontext für die Druckerauswahl zurück.
'    cdlPDReturnDefault = 1024 ' (&H400) ' Gibt den Namen des Standarddruckers zurück.
'    cdlPDHelpButton = 2048    ' (&H800) ' Das Dialogfeld zeigt die Schaltfläche 'Hilfe' an.
'    cdlPDUseDevModeCopies = 262144   ' (&H40000)  ' Legt die Unterstützung für mehrere Exemplare fest.
'    cdlPDDisablePrintToFile = 524288 ' (&H80000)  ' Deaktiviert das Kontrollkästchen 'Ausdruck in Datei'.
'    cdlPDHidePrintToFile = 1048576   ' (&H100000) ' Das Kontrollkästchen 'Ausdruck in Datei' wird nicht angezeigt.
'End Enum

Private m_IsInitialized As Boolean
Private m_PaperSizes() As PaperSize

Public Sub Init()
    InitPaperSizes
    m_IsInitialized = True
End Sub

Public Property Get PaperSizes_Item(ByVal Index As Long) As PaperSize
    PaperSizes_Item = m_PaperSizes(Index)
End Property

Private Function New_PaperSize(ByVal Index As Long, ByVal Name As String, ByVal Width_inch100 As Long, ByVal Height_inch100 As Long) As PaperSize
    With New_PaperSize: .RawKind = Index: .PaperName = Name: .Kind = .RawKind: .Width = Width_inch100: .Height = Height_inch100: End With
End Function

Private Function InitPaperSizes()
    ReDim m_PaperSizes(0 To 256)
    m_PaperSizes(1) = New_PaperSize(1, "Letter", 850, 1100)
    m_PaperSizes(2) = New_PaperSize(2, "LetterSmall", 850, 1100)
    m_PaperSizes(3) = New_PaperSize(3, "Tabloid", 1100, 1700)
    m_PaperSizes(4) = New_PaperSize(4, "Ledger", 1700, 1100)
    m_PaperSizes(5) = New_PaperSize(5, "Legal", 850, 1400)
    m_PaperSizes(6) = New_PaperSize(6, "Statement", 550, 850)
    m_PaperSizes(7) = New_PaperSize(7, "Executive", 725, 1050)
    m_PaperSizes(8) = New_PaperSize(8, "DIN-A3", 1169, 1654)
    m_PaperSizes(9) = New_PaperSize(9, "DIN-A4", 827, 1169)
    m_PaperSizes(10) = New_PaperSize(10, "DIN-A4Small", 827, 1169)
    m_PaperSizes(11) = New_PaperSize(11, "DIN-A5", 583, 827)
    m_PaperSizes(12) = New_PaperSize(12, "DIN-B4", 1012, 1433)
    m_PaperSizes(13) = New_PaperSize(13, "DIN-B5", 717, 1012)
    m_PaperSizes(14) = New_PaperSize(14, "Folio", 850, 1300)
    m_PaperSizes(15) = New_PaperSize(15, "Quarto", 846, 1083)
    m_PaperSizes(16) = New_PaperSize(16, "Standard10x14", 1000, 1400)
    m_PaperSizes(17) = New_PaperSize(17, "Standard11x17", 1100, 1700)
    m_PaperSizes(18) = New_PaperSize(18, "Note", 850, 1100)
    m_PaperSizes(19) = New_PaperSize(19, "Number9Envelope", 387, 887)
    m_PaperSizes(20) = New_PaperSize(20, "Number10Envelope", 412, 950)
    m_PaperSizes(21) = New_PaperSize(21, "Number11Envelope", 450, 1037)
    m_PaperSizes(22) = New_PaperSize(22, "Number12Envelope", 475, 1100)
    m_PaperSizes(23) = New_PaperSize(23, "Number14Envelope", 500, 1150)
    m_PaperSizes(24) = New_PaperSize(24, "CSheet", 1700, 2200)
    m_PaperSizes(25) = New_PaperSize(25, "DSheet", 2200, 3400)
    m_PaperSizes(26) = New_PaperSize(26, "ESheet", 3400, 4400)
    m_PaperSizes(27) = New_PaperSize(27, "DLEnvelope", 433, 866)
    m_PaperSizes(28) = New_PaperSize(28, "DIN-C5Envelope", 638, 902)
    m_PaperSizes(29) = New_PaperSize(29, "DIN-C3Envelope", 1276, 1803)
    m_PaperSizes(30) = New_PaperSize(30, "DIN-C4Envelope", 902, 1276)
    m_PaperSizes(31) = New_PaperSize(31, "DIN-C6Envelope", 449, 638)
    m_PaperSizes(32) = New_PaperSize(32, "DIN-C65Envelope", 449, 902)
    m_PaperSizes(33) = New_PaperSize(33, "DIN-B4Envelope", 984, 1390)
    m_PaperSizes(34) = New_PaperSize(34, "DIN-B5Envelope", 693, 984)
    m_PaperSizes(35) = New_PaperSize(35, "DIN-B6Envelope", 693, 492)
    m_PaperSizes(36) = New_PaperSize(36, "ItalyEnvelope", 433, 906)
    m_PaperSizes(37) = New_PaperSize(37, "MonarchEnvelope", 387, 750)
    m_PaperSizes(38) = New_PaperSize(38, "PersonalEnvelope", 362, 650)
    m_PaperSizes(39) = New_PaperSize(39, "USStandardFanfold", 1487, 1100)
    m_PaperSizes(40) = New_PaperSize(40, "GermanStandardFanfold", 850, 1200)
    m_PaperSizes(41) = New_PaperSize(41, "GermanLegalFanfold", 850, 1300)
    m_PaperSizes(42) = New_PaperSize(42, "DIN-IsoB4", 984, 1390)
    m_PaperSizes(43) = New_PaperSize(43, "JapanesePostcard", 394, 583)
    m_PaperSizes(44) = New_PaperSize(44, "Standard9x11", 900, 1100)
    m_PaperSizes(45) = New_PaperSize(45, "Standard10x11", 1000, 1100)
    m_PaperSizes(46) = New_PaperSize(46, "Standard15x11", 1500, 1100)
    m_PaperSizes(47) = New_PaperSize(47, "InviteEnvelope", 866, 866)
    '48
    '49
    m_PaperSizes(50) = New_PaperSize(50, "LetterExtra", 950, 1200)
    m_PaperSizes(51) = New_PaperSize(51, "LegalExtra", 950, 1500)
    m_PaperSizes(53) = New_PaperSize(53, "DIN-A4Extra", 927, 1269)
    m_PaperSizes(54) = New_PaperSize(54, "LetterTransverse", 850, 1100)
    m_PaperSizes(55) = New_PaperSize(55, "DIN-A4Transverse", 827, 1169)
    m_PaperSizes(56) = New_PaperSize(56, "LetterExtraTransverse", 950, 1200)
    m_PaperSizes(57) = New_PaperSize(57, "DIN-APlus", 894, 1402)
    m_PaperSizes(58) = New_PaperSize(58, "DIN-BPlus", 1201, 1917)
    m_PaperSizes(59) = New_PaperSize(59, "LetterPlus", 850, 1269)
    m_PaperSizes(60) = New_PaperSize(60, "DIN-A4Plus", 827, 1299)
    m_PaperSizes(61) = New_PaperSize(61, "DIN-A5Transverse", 583, 827)
    m_PaperSizes(62) = New_PaperSize(62, "DIN-B5Transverse", 717, 1012)
    m_PaperSizes(63) = New_PaperSize(63, "DIN-A3Extra", 1268, 1752)
    m_PaperSizes(64) = New_PaperSize(64, "DIN-A5Extra", 685, 925)
    m_PaperSizes(65) = New_PaperSize(65, "DIN-B5Extra", 791, 1087)
    m_PaperSizes(66) = New_PaperSize(66, "DIN-A2", 1654, 2339)
    m_PaperSizes(67) = New_PaperSize(67, "DIN-A3Transverse", 1169, 1654)
    m_PaperSizes(68) = New_PaperSize(68, "DIN-A3ExtraTransverse", 1268, 1752)
    m_PaperSizes(69) = New_PaperSize(69, "JapaneseDoublePostcard", 787, 583)
    m_PaperSizes(70) = New_PaperSize(70, "DIN-A6", 413, 583)
    m_PaperSizes(71) = New_PaperSize(71, "JapaneseEnvelopeKakuNumber2", 945, 1307)
    m_PaperSizes(72) = New_PaperSize(72, "JapaneseEnvelopeKakuNumber3", 850, 1091)
    m_PaperSizes(73) = New_PaperSize(73, "JapaneseEnvelopeChouNumber3", 472, 925)
    m_PaperSizes(74) = New_PaperSize(74, "JapaneseEnvelopeChouNumber4", 354, 807)
    m_PaperSizes(75) = New_PaperSize(75, "LetterRotated", 1100, 850)
    m_PaperSizes(76) = New_PaperSize(76, "DIN-A3Rotated", 1654, 1169)
    m_PaperSizes(77) = New_PaperSize(77, "DIN-A4Rotated", 1169, 827)
    m_PaperSizes(78) = New_PaperSize(78, "DIN-A5Rotated", 827, 583)
    m_PaperSizes(79) = New_PaperSize(79, "DIN-B4JisRotated", 1433, 1012)
    m_PaperSizes(80) = New_PaperSize(80, "DIN-B5JisRotated", 1012, 717)
    m_PaperSizes(81) = New_PaperSize(81, "JapanesePostcardRotated", 583, 394)
    m_PaperSizes(82) = New_PaperSize(82, "JapaneseDoublePostcardRotated", 583, 787)
    m_PaperSizes(83) = New_PaperSize(83, "DIN-A6Rotated", 583, 413)
    m_PaperSizes(84) = New_PaperSize(84, "JapaneseEnvelopeKakuNumber2Rotated", 1307, 945)
    m_PaperSizes(85) = New_PaperSize(85, "JapaneseEnvelopeKakuNumber3Rotated", 1091, 850)
    m_PaperSizes(86) = New_PaperSize(86, "JapaneseEnvelopeChouNumber3Rotated", 925, 472)
    m_PaperSizes(87) = New_PaperSize(87, "JapaneseEnvelopeChouNumber4Rotated", 807, 354)
    m_PaperSizes(88) = New_PaperSize(88, "DIN-B6Jis", 504, 717)
    m_PaperSizes(89) = New_PaperSize(89, "DIN-B6JisRotated", 717, 504)
    m_PaperSizes(90) = New_PaperSize(90, "Standard12x11", 1200, 1100)
    m_PaperSizes(91) = New_PaperSize(91, "JapaneseEnvelopeYouNumber4", 413, 925)
    m_PaperSizes(92) = New_PaperSize(92, "JapaneseEnvelopeYouNumber4Rotated", 925, 413)
    m_PaperSizes(96) = New_PaperSize(96, "PrcEnvelopeNumber1", 402, 650)
    m_PaperSizes(98) = New_PaperSize(98, "PrcEnvelopeNumber3", 492, 693)
    m_PaperSizes(99) = New_PaperSize(99, "PrcEnvelopeNumber4", 433, 819)
    m_PaperSizes(100) = New_PaperSize(100, "PrcEnvelopeNumber5", 433, 866)
    m_PaperSizes(101) = New_PaperSize(101, "PrcEnvelopeNumber6", 472, 906)
    m_PaperSizes(102) = New_PaperSize(102, "PrcEnvelopeNumber7", 630, 906)
    m_PaperSizes(103) = New_PaperSize(103, "PrcEnvelopeNumber8", 472, 1217)
    m_PaperSizes(104) = New_PaperSize(104, "PrcEnvelopeNumber9", 902, 1276)
    m_PaperSizes(105) = New_PaperSize(105, "PrcEnvelopeNumber10", 1276, 1803)
    m_PaperSizes(109) = New_PaperSize(109, "PrcEnvelopeNumber1Rotated", 650, 402)
    m_PaperSizes(111) = New_PaperSize(111, "PrcEnvelopeNumber3Rotated", 693, 492)
    m_PaperSizes(112) = New_PaperSize(112, "PrcEnvelopeNumber4Rotated", 819, 433)
    m_PaperSizes(113) = New_PaperSize(113, "PrcEnvelopeNumber5Rotated", 866, 433)
    m_PaperSizes(114) = New_PaperSize(114, "PrcEnvelopeNumber6Rotated", 906, 472)
    m_PaperSizes(115) = New_PaperSize(115, "PrcEnvelopeNumber7Rotated", 906, 630)
    m_PaperSizes(116) = New_PaperSize(116, "PrcEnvelopeNumber8Rotated", 1217, 472)
    m_PaperSizes(117) = New_PaperSize(117, "PrcEnvelopeNumber9Rotated", 1276, 902)
    m_PaperSizes(256) = New_PaperSize(256, "Custom", 827, 1169)
    m_IsInitialized = True
End Function

Public Function PaperKind_ToStr(this As PaperKind) As String
    If Not m_IsInitialized Then Init
    If this < 0 Or 256 < this Then Exit Function
    PaperKind_ToStr = m_PaperSizes(this).PaperName
End Function

Public Function PaperKind_Parse(ByVal s As String) As PaperKind
    Dim pk As PaperKind
    Select Case s
    Case "Custom":                            pk = PaperKind.Custom
    Case "Letter":                            pk = PaperKind.Letter
    Case "LetterSmall":                       pk = PaperKind.LetterSmall
    Case "Tabloid":                           pk = PaperKind.Tabloid
    Case "Ledger":                            pk = PaperKind.Ledger
    Case "Legal":                             pk = PaperKind.Legal
    Case "Statement":                         pk = PaperKind.Statement
    Case "Executive":                         pk = PaperKind.Executive
    Case "DIN-A3":                            pk = PaperKind.DIN_A3
    Case "DIN-A4":                            pk = PaperKind.DIN_A4
    Case "DIN-A4Small":                       pk = PaperKind.DIN_A4Small
    Case "DIN-A5":                            pk = PaperKind.DIN_A5
    Case "DIN-B4":                            pk = PaperKind.DIN_B4
    Case "DIN-B5":                            pk = PaperKind.DIN_B5
    Case "Folio":                             pk = PaperKind.Folio
    Case "Quarto":                            pk = PaperKind.Quarto
    Case "Standard10x14":                     pk = PaperKind.Standard10x14
    Case "Standard11x17":                     pk = PaperKind.Standard11x17
    Case "Note":                              pk = PaperKind.Note
    Case "Number9Envelope":                   pk = PaperKind.Number9Envelope
    Case "Number10Envelope":                  pk = PaperKind.Number10Envelope
    Case "Number11Envelope":                  pk = PaperKind.Number11Envelope
    Case "Number12Envelope":                  pk = PaperKind.Number12Envelope
    Case "Number14Envelope":                  pk = PaperKind.Number14Envelope
    Case "CSheet":                            pk = PaperKind.CSheet
    Case "DSheet":                            pk = PaperKind.DSheet
    Case "ESheet":                            pk = PaperKind.ESheet
    Case "DLEnvelope":                        pk = PaperKind.DLEnvelope
    Case "DIN-C5Envelope":                    pk = PaperKind.DIN_C5Envelope
    Case "DIN-C3Envelope":                    pk = PaperKind.DIN_C3Envelope
    Case "DIN-C4Envelope":                    pk = PaperKind.DIN_C4Envelope
    Case "DIN-C6Envelope":                    pk = PaperKind.DIN_C6Envelope
    Case "DIN-C65Envelope":                   pk = PaperKind.DIN_C65Envelope
    Case "DIN-B4Envelope":                    pk = PaperKind.DIN_B4Envelope
    Case "DIN-B5Envelope":                    pk = PaperKind.DIN_B5Envelope
    Case "DIN-B6Envelope":                    pk = PaperKind.DIN_B6Envelope
    Case "ItalyEnvelope":                     pk = PaperKind.ItalyEnvelope
    Case "MonarchEnvelope":                   pk = PaperKind.MonarchEnvelope
    Case "PersonalEnvelope":                  pk = PaperKind.PersonalEnvelope
    Case "USStandardFanfold":                 pk = PaperKind.USStandardFanfold
    Case "GermanStandardFanfold":             pk = PaperKind.GermanStandardFanfold
    Case "GermanLegalFanfold":                pk = PaperKind.GermanLegalFanfold
    Case "DIN-IsoB4":                         pk = PaperKind.DIN_IsoB4
    Case "JapanesePostcard":                  pk = PaperKind.JapanesePostcard
    Case "Standard9x11":                      pk = PaperKind.Standard9x11
    Case "Standard10x11":                     pk = PaperKind.Standard10x11
    Case "Standard15x11":                     pk = PaperKind.Standard15x11
    Case "InviteEnvelope":                    pk = PaperKind.InviteEnvelope
    '
    '
    Case "LetterExtra":                       pk = PaperKind.LetterExtra
    Case "LegalExtra":                        pk = PaperKind.LegalExtra
    Case "TabloidExtra":                      pk = PaperKind.TabloidExtra
    Case "DIN_A4Extra":                       pk = PaperKind.DIN_A4Extra
    Case "LetterTransverse":                  pk = PaperKind.LetterTransverse
    Case "DIN_A4Transverse":                  pk = PaperKind.DIN_A4Transverse
    Case "LetterExtraTransverse":             pk = PaperKind.LetterExtraTransverse
    Case "DIN-APlus":                         pk = PaperKind.APlus
    Case "DIN-BPlus":                         pk = PaperKind.BPlus
    Case "LetterPlus":                        pk = PaperKind.LetterPlus
    Case "DIN-A4Plus":                        pk = PaperKind.DIN_A4Plus
    Case "DIN-A5Transverse":                  pk = PaperKind.DIN_A5Transverse
    Case "DIN-B5Transverse":                  pk = PaperKind.DIN_B5Transverse
    Case "DIN-A3Extra":                       pk = PaperKind.DIN_A3Extra
    Case "DIN-A5Extra":                       pk = PaperKind.DIN_A5Extra
    Case "DIN-B5Extra":                       pk = PaperKind.DIN_B5Extra
    Case "DIN-A2":                            pk = PaperKind.DIN_A2
    Case "DIN-A3Transverse":                  pk = PaperKind.DIN_A3Transverse
    Case "DIN-A3ExtraTransverse":             pk = PaperKind.DIN_A3ExtraTransverse
    Case "JapaneseDoublePostcard":            pk = PaperKind.JapaneseDoublePostcard
    Case "DIN-A6":                            pk = PaperKind.DIN_A6
    Case "JapaneseEnvelopeKakuNumber2":       pk = PaperKind.JapaneseEnvelopeKakuNumber2
    Case "JapaneseEnvelopeKakuNumber3":       pk = PaperKind.JapaneseEnvelopeKakuNumber3
    Case "JapaneseEnvelopeChouNumber3":       pk = PaperKind.JapaneseEnvelopeChouNumber3
    Case "JapaneseEnvelopeChouNumber4":       pk = PaperKind.JapaneseEnvelopeChouNumber4
    Case "LetterRotated":                     pk = PaperKind.LetterRotated
    Case "DIN-A3Rotated":                     pk = PaperKind.DIN_A3Rotated
    Case "DIN-A4Rotated":                     pk = PaperKind.DIN_A4Rotated
    Case "DIN-A5Rotated":                     pk = PaperKind.DIN_A5Rotated
    Case "DIN-B4JisRotated":                  pk = PaperKind.DIN_B4JisRotated
    Case "DIN-B5JisRotated":                  pk = PaperKind.DIN_B5JisRotated
    Case "JapanesePostcardRotated":           pk = PaperKind.JapanesePostcardRotated
    Case "JapaneseDoublePostcardRotated":     pk = PaperKind.JapaneseDoublePostcardRotated
    Case "DIN-A6Rotated":                     pk = PaperKind.DIN_A6Rotated
    Case "JapaneseEnvelopeKakuNumber2Rotated": pk = PaperKind.JapaneseEnvelopeKakuNumber2Rotated
    Case "JapaneseEnvelopeKakuNumber3Rotated": pk = PaperKind.JapaneseEnvelopeKakuNumber3Rotated
    Case "JapaneseEnvelopeChouNumber3Rotated": pk = PaperKind.JapaneseEnvelopeChouNumber3Rotated
    Case "JapaneseEnvelopeChouNumber4Rotated": pk = PaperKind.JapaneseEnvelopeChouNumber4Rotated
    Case "DIN-B6Jis":                         pk = PaperKind.DIN_B6Jis
    Case "DIN-B6JisRotated":                  pk = PaperKind.DIN_B6JisRotated
    Case "Standard12x11":                     pk = PaperKind.Standard12x11
    Case "JapaneseEnvelopeYouNumber4":        pk = PaperKind.JapaneseEnvelopeYouNumber4
    Case "JapaneseEnvelopeYouNumber4Rotated": pk = PaperKind.JapaneseEnvelopeYouNumber4Rotated
    Case "Prc16K":                            pk = PaperKind.Prc16K
    Case "Prc32K":                            pk = PaperKind.Prc32K
    Case "Prc32KBig":                         pk = PaperKind.Prc32KBig
    Case "PrcEnvelopeNumber1":                pk = PaperKind.PrcEnvelopeNumber1
    Case "PrcEnvelopeNumber2":                pk = PaperKind.PrcEnvelopeNumber2
    Case "PrcEnvelopeNumber3":                pk = PaperKind.PrcEnvelopeNumber3
    Case "PrcEnvelopeNumber4":                pk = PaperKind.PrcEnvelopeNumber4
    Case "PrcEnvelopeNumber5":                pk = PaperKind.PrcEnvelopeNumber5
    Case "PrcEnvelopeNumber6":                pk = PaperKind.PrcEnvelopeNumber6
    Case "PrcEnvelopeNumber7":                pk = PaperKind.PrcEnvelopeNumber7
    Case "PrcEnvelopeNumber8":                pk = PaperKind.PrcEnvelopeNumber8
    Case "PrcEnvelopeNumber9":                pk = PaperKind.PrcEnvelopeNumber9
    Case "PrcEnvelopeNumber10":               pk = PaperKind.PrcEnvelopeNumber10
    Case "Prc16KRotated":                     pk = PaperKind.Prc16KRotated
    Case "Prc32KRotated":                     pk = PaperKind.Prc32KRotated
    Case "Prc32KBigRotated":                  pk = PaperKind.Prc32KBigRotated
    Case "PrcEnvelopeNumber1Rotated":         pk = PaperKind.PrcEnvelopeNumber1Rotated
    Case "PrcEnvelopeNumber2Rotated":         pk = PaperKind.PrcEnvelopeNumber2Rotated
    Case "PrcEnvelopeNumber3Rotated":         pk = PaperKind.PrcEnvelopeNumber3Rotated
    Case "PrcEnvelopeNumber4Rotated":         pk = PaperKind.PrcEnvelopeNumber4Rotated
    Case "PrcEnvelopeNumber5Rotated":         pk = PaperKind.PrcEnvelopeNumber5Rotated
    Case "PrcEnvelopeNumber6Rotated":         pk = PaperKind.PrcEnvelopeNumber6Rotated
    Case "PrcEnvelopeNumber7Rotated":         pk = PaperKind.PrcEnvelopeNumber7Rotated
    Case "PrcEnvelopeNumber8Rotated":         pk = PaperKind.PrcEnvelopeNumber8Rotated
    Case "PrcEnvelopeNumber9Rotated":         pk = PaperKind.PrcEnvelopeNumber9Rotated
    Case "PrcEnvelopeNumber10Rotated":        pk = PaperKind.PrcEnvelopeNumber10Rotated
    End Select
    PaperKind_Parse = pk
End Function

Public Sub PaperKind_ToListBox(aLBCB) 'aLBCB As ListBox Or As ComboBox
    If Not m_IsInitialized Then Init
    aLBCB.Clear
    Dim i As Long, s As String
    For i = 0 To UBound(m_PaperSizes)
        s = PaperKind_ToStr(i)
        If Len(s) Then aLBCB.AddItem s
    Next
End Sub

Public Function PaperSize_ToStr(this As PaperSize) As String
    Dim s As String: s = "[PaperSize"
    With this
        s = s & " Kind=" & .PaperName
        s = s & " Height=" & .Height
        s = s & " Width=" & .Width
    End With
    PaperSize_ToStr = s & "]"
End Function

Public Function PaperOrientation_ToStr(ByVal po As PrinterOrientationConstants) As String
    Dim s As String
    Select Case po
    Case PrinterOrientationConstants.cdlPortrait:  s = "Hochformat" '"Portrait"
    Case PrinterOrientationConstants.cdlLandscape: s = "Querformat" '"Landscape"
    Case Else: s = CStr(po)
    End Select
    PaperOrientation_ToStr = s
End Function

Public Function PaperSource_ToStr(ByVal psk As PaperSourceKind) As String
    Dim s As String
    Select Case psk
    Case 1:    s = "Upper"
    Case 2:    s = "Lower"
    Case 3:    s = "Middle"
    Case 4:    s = "Manual"
    Case 5:    s = "Envelope"
    Case 6:    s = "ManualFeed"
    Case 7:    s = "AutomaticFeed"
    Case 8:    s = "TractorFeed"
    Case 9:    s = "SmallFormat"
    Case 10:   s = "LargeFormat"
    Case 11:   s = "LargeCapacity"
    Case 14:   s = "Cassette"
    Case 15:   s = "FormSource"
    Case 257:  s = "Custom"
    Case Else: s = CStr(psk)
    End Select
    PaperSource_ToStr = s
End Function

Public Function PrinterConstants_ToStr(prc As PrinterConstants) As String
    Dim s As String
    If prc And cdlPDAllPages Then s = s & IIf(Len(s) > 0, ", ", "") & "AllPages"
    If prc And cdlPDSelection Then s = s & IIf(Len(s) > 0, ", ", "") & "Selection"
    If prc And cdlPDPageNums Then s = s & IIf(Len(s) > 0, ", ", "") & "PageNums"
    If prc And cdlPDNoSelection Then s = s & IIf(Len(s) > 0, ", ", "") & "NoSelection"
    If prc And cdlPDNoPageNums Then s = s & IIf(Len(s) > 0, ", ", "") & "NoPageNums"
    If prc And cdlPDCollate Then s = s & IIf(Len(s) > 0, ", ", "") & "Collate"
    If prc And cdlPDPrintToFile Then s = s & IIf(Len(s) > 0, ", ", "") & "PrintToFile"
    If prc And cdlPDPrintSetup Then s = s & IIf(Len(s) > 0, ", ", "") & "PrintSetup"
    If prc And cdlPDNoWarning Then s = s & IIf(Len(s) > 0, ", ", "") & "NoWarning"
    If prc And cdlPDReturnDC Then s = s & IIf(Len(s) > 0, ", ", "") & "ReturnDC"
    If prc And cdlPDReturnIC Then s = s & IIf(Len(s) > 0, ", ", "") & "ReturnIC"
    If prc And cdlPDReturnDefault Then s = s & IIf(Len(s) > 0, ", ", "") & "ReturnDefault"
    If prc And cdlPDHelpButton Then s = s & IIf(Len(s) > 0, ", ", "") & "HelpButton"
    If prc And cdlPDUseDevModeCopies Then s = s & IIf(Len(s) > 0, ", ", "") & "UseDevModeCopies"
    If prc And cdlPDDisablePrintToFile Then s = s & IIf(Len(s) > 0, ", ", "") & "DisablePrintToFile"
    If prc And cdlPDHidePrintToFile Then s = s & IIf(Len(s) > 0, ", ", "") & "HidePrintToFile"
    PrinterConstants_ToStr = s
End Function

