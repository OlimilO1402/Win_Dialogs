VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame2 
      Caption         =   "Printer"
      Height          =   2895
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   3975
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Label14"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Label13"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Label12"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Label8"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CommonDialog.ShowPrinter"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Label7"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   480
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy from CommonDialog to Printer"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CommonDialog.ShowPrinter"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PaperSourceKind
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

Private Sub Form_Load()
    Label1.Caption = "Printer:     "
    Label2.Caption = "Copies:      "
    Label3.Caption = "PaperBin:    "
    Label4.Caption = "PaperSize:   "
    Label5.Caption = "CopyToFile:  "
    Label6.Caption = "Orientation: "
    Label7.Caption = "Page min-max:"
    
    Label8.Caption = Label1.Caption
    Label9.Caption = Label2.Caption
    Label10.Caption = Label3.Caption
    Label11.Caption = Label4.Caption
    Label12.Caption = Label5.Caption
    Label13.Caption = Label6.Caption
    Label14.Caption = Label7.Caption
    UpdateView
End Sub

Private Sub Command1_Click()
    'Show the Printer-dialog
Try: On Error GoTo Catch
    With CommonDialog1
        .Flags = .Flags Or PrinterConstants.cdlPDPrintSetup
        .Flags = .Flags Or PrinterConstants.cdlPDUseDevModeCopies
        .Min = 1
        .Max = 20
        .CancelError = True
        .ShowPrinter
    End With
    UpdateView
    Exit Sub
Catch:
    Select Case Err.Number
    Case 32755: On Error GoTo 0: Exit Sub
    Case Else: MsgBox Err.Description
    End Select
End Sub

Private Sub Command2_Click()
    'Copy data from Commondialog to the Printer-object:
    'Printer.DeviceName = CommonDialog1.???
    Printer.Copies = CommonDialog1.Copies
    'Printer.PaperSize = CommonDialog1.???
    Printer.Orientation = CommonDialog1.Orientation
    'Printer.Min = CommonDialog1.FromPage
    'Printer.Max = CommonDialog1.ToPage
    UpdateView
End Sub

Sub UpdateView()
    With CommonDialog1
        Label1.Caption = "Printer:   " '& .???
        Label2.Caption = "Copies:    " & .Copies
        Label3.Caption = "PaperBin:  " '& .???
        Label4.Caption = "PaperSize: " '& ???
        Label5.Caption = "CopyToFile:  " & ((.Flags And PrinterConstants.cdlPDPrintToFile) = PrinterConstants.cdlPDPrintToFile)
        Label6.Caption = "Orientation: " & PaperOrientation_ToStr(.Orientation)
        Label7.Caption = "Page min-max: " & .FromPage & "-" & .ToPage
    End With
    With Printer
        Label8.Caption = "Printer:   " & .DeviceName
        Label9.Caption = "Copies:    " & .Copies
        Label10.Caption = "PaperBin:  " & PaperSource_ToStr(.PaperBin)
        Label11.Caption = "PaperSize: " & PaperSize_ToStr(.PaperSize)
        Label12.Caption = "CopyToFile:  "
        Label13.Caption = "Orientation: " & PaperOrientation_ToStr(.Orientation)
        Label14.Caption = "Page min-max: " ' & .???
    End With
End Sub

Private Function PaperSize_ToStr(ByVal ps As Long) As String
    Dim s As String
    Select Case ps
    Case 1:    s = "Letter"           ' Letter paper (8.5 in.by 11 in.).
    Case 2:    s = "LetterSmall"      ' Letter small paper (8.5 in.by 11 in.)."
    Case 3:    s = "Tabloid"          ' Tabloid paper (11 in.by 17 in.)."
    Case 4:    s = "Ledger"           ' Ledger paper (17 in.by 11 in.)."
    Case 5:    s = "Legal"            ' Legal paper (8.5 in.by 14 in.)."
    Case 6:    s = "Statement"        ' Statement paper (5.5 in.by 8.5 in.)."
    Case 7:    s = "Executive"        ' Executive paper (7.25 in.by 10.5 in.)."
    Case 8:    s = "DIN-A3"           ' A3 paper (297 mm by 420 mm)."
    Case 9:    s = "DIN-A4"           ' A4 (210 x 297 mm)."
    Case 10:   s = "DIN-A4Small"      ' A4 klein (210 x 297 mm)."
    Case 11:   s = "DIN-A5"           ' A5 (148 x 210 mm)."
    Case 12:   s = "DIN-B4"           ' B4 (250 x 353 mm)."
    Case 13:   s = "DIN-B5"           ' B5 (176 x 250 mm)."
    Case 14:   s = "Folio"            ' Folio paper (8.5 in.by 13 in.)."
    Case 15:   s = "Quarto"           ' Quarto (215 x 275 mm)."
    Case 16:   s = "Standard10x14"    ' Standard paper (10 in.by 14 in.)."
    Case 17:   s = "Standard11x17"    ' Standard paper (11 in.by 17 in.)."
    Case 18:   s = "Note"             ' Note paper (8.5 in.by 11 in.)."
    Case 19:   s = "Number9Envelope"  ' #9 envelope (3.875 in.by 8.875 in.)."
    Case 20:   s = "Number10Envelope" ' #10 envelope (4.125 in.by 9.5 in.)."
    'Case ...
    Case 256:  s = "Custom"
    Case Else: s = CStr(ps)
    End Select
    PaperSize_ToStr = s
End Function

Private Function PaperOrientation_ToStr(ByVal po As PrinterOrientationConstants) As String
    Dim s As String
    Select Case po
    Case PrinterOrientationConstants.cdlPortrait:  s = "Hochformat"
    Case PrinterOrientationConstants.cdlLandscape: s = "Querformat"
    Case Else: s = CStr(po)
    End Select
    PaperOrientation_ToStr = s
End Function

Private Function PaperSource_ToStr(ByVal psk As PaperSourceKind) As String
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
