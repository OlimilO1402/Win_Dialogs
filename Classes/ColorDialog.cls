VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class ColorDialog
'          Inherits System.Windows.Forms.CommonDialog
'     Member von: System.Windows.Forms
'
'Zusammenfassung:
' Stellt ein Standarddialogfeld dar, in dem die verfügbaren Farben angezeigt werden
' wie auch Steuerelemente, mit denen Benutzer benutzerdefinierte Farben definieren
' können.
Option Explicit
'Implements CommonDialog
Private Const CC_RGBINIT              As Long = &H1   ' Vorgabe einer Standard-Farbe
Private Const CC_FULLOPEN             As Long = &H2   ' Vollauswahl aller Farben anzeigen
Private Const CC_PREVENTFULLOPEN      As Long = &H4   ' Deaktiviert den Button zum Öffnen der Dialogbox-Erweiterung
Private Const CC_SHOWHELP             As Long = &H8   ' Hilfe-Button anzeigen
Private Const CC_ENABLEHOOK           As Long = &H10  ' Nachrichten können "abgefangen" werden
Private Const CC_ENABLETEMPLATE       As Long = &H20  ' Dialogbox Template
Private Const CC_ENABLETEMPLATEHANDLE As Long = &H40  ' Benutzt Template, ignoriert aber den Template-Namen
Private Const CC_SOLIDCOLOR           As Long = &H80  ' nur Grundfarben auswählbar
Private Const CC_ANYCOLOR             As Long = &H100 ' Anwender kann alle Farben wählen

'possible errors in Dialog ChooseColor
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
Private Const CDERR_DIALOGFAILURE   As Long = &HFFFF&

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
'typedef struct tagCHOOSECOLORW {
'  DWORD        lStructSize;
'  HWND         hwndOwner;
'  HWND         hInstance;
'  COLORREF     rgbResult;
'  COLORREF     *lpCustColors;
'  DWORD        Flags;
'  LPARAM       lCustData;
'  LPCCHOOKPROC lpfnHook;
'  LPCWSTR      lpTemplateName;
'  LPEDITMENU   lpEditInfo;       ' Käse dat gibbet nich!!
'} CHOOSECOLORW, *LPCHOOSECOLORW;

Private Type ChooseColorW     '      x86 '     x64
    lStructSize    As Long    '        4 '       4 + 4pads
    hwndOwner      As LongPtr '        4 '       8
    hInstance      As LongPtr '        4 '       8
    rgbResult      As Long    '        4 '       4 + 4pads
    lpCustColors   As LongPtr '        4 '       8
    Flags          As Long    '        4 '       4 + 4pads
    lCustData      As LongPtr '        4 '       8
    lpfnHook       As LongPtr '        4 '       8
    lpTemplateName As LongPtr 'String  4 '       8
End Type                      '  Sum: 36 ' Sum: 72
#If VBA7 Then
    Private Declare PtrSafe Function ChooseColorW Lib "comdlg32" (pChooseColor As ChooseColorW) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
    Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long
#Else
    Private Declare Function ChooseColorW Lib "comdlg32" (pChooseColor As ChooseColorW) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
#End If

Private m_Flags As Long
Private Const MaxCustomColors As Long = 16
Private m_CustomColors() As Long 'mehr als 16 sind im Dialog gar nicht möglich
'wenn der Cancel-Button gedrückt wurde dann soll auch die alte Farbtabelle wiederhergestellt werden
Private m_Color As Long
Private m_TemplateName As String

Private Sub Class_Initialize()
    '
    Call Reset
End Sub
'Public Sub New() 'keine Parameter, keine externe Constructor Funktion erforderlich

'Public Overridable Property AllowFullOpen() As Boolean
Public Property Get AllowFullOpen() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld benutzerdefinierte Farben definiert
    ' werden können, oder legt diesen fest.
    AllowFullOpen = Not (m_Flags And CC_PREVENTFULLOPEN)
End Property
Public Property Let AllowFullOpen(ByVal Value As Boolean)
    m_Flags = m_Flags Or CC_PREVENTFULLOPEN
    If Value Then
        m_Flags = m_Flags Xor CC_PREVENTFULLOPEN
    End If
End Property
'Public Overridable Property FullOpen() As Boolean
Public Property Get FullOpen() As Boolean
    ' Ruft einen Wert ab, der angibt, ob die Steuerelemente für das Erstellen
    ' benutzerdefinierter Farben beim Öffnen des Dialogfelds angezeigt werden,
    ' oder legt diesen fest.
    FullOpen = Not (m_Flags And CC_FULLOPEN)
End Property
Public Property Let FullOpen(ByVal Value As Boolean)
    m_Flags = m_Flags Or CC_FULLOPEN
    If Not Value Then
        m_Flags = m_Flags Xor CC_FULLOPEN
    End If
End Property

'Public Overridable Property AnyColor() As Boolean
Public Property Get AnyColor() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld bei den Grundfarben alle verfügbaren
    ' Farben angezeigt werden, oder legt diesen fest.
    AnyColor = (m_Flags And CC_ANYCOLOR)
End Property
Public Property Let AnyColor(ByVal Value As Boolean)
    m_Flags = m_Flags Or CC_ANYCOLOR
    If Not Value Then
        m_Flags = m_Flags Xor CC_ANYCOLOR
    End If
End Property

'Public Property Color() As System.Drawing.Color
Public Property Get Color() As Long 'System.Drawing.Color
    ' Ruft die von den Benutzern ausgewählte Farbe ab oder legt diese fest.
    Color = m_Color
End Property
Public Property Let Color(ByVal Value As Long)  'System.Drawing.Color
    m_Color = Value
End Property

'Public Property CustomColors() As Integer()
Public Property Get CustomColors(ByVal Index As Long) As Long 'Integer()
    ' Ruft den im Dialogfeld angezeigten Satz benutzerdefinierter Farben ab
    ' oder legt diesen fest.
    If Index > 16 Then Exit Property
    CustomColors = m_CustomColors(Index)
End Property
Public Property Let CustomColors(ByVal Index As Long, ByVal Value As Long) 'Integer()
    If Index > 16 Then Exit Property
    m_CustomColors(Index) = Value
End Property

'Public Overrides Sub Reset()
Public Sub Reset()
    ' Setzt alle Optionen auf die Standardwerte, die zuletzt ausgewählte Farbe auf
    ' schwarz und die benutzerdefinierten Farben auf die Standardwerte zurück.
    ' welche Standardwerte der UD Farben?
    m_Flags = m_Flags Or CC_RGBINIT
    m_Flags = m_Flags Or CC_ANYCOLOR
    m_Flags = m_Flags Or CC_FULLOPEN
    m_Flags = m_Flags Or CC_PREVENTFULLOPEN
    m_Color = vbBlack '&H0& 'Schwarz '
    ReDim m_CustomColors(0 To MaxCustomColors - 1) 'As Long 'mehr als 16 sind im Dialog gar nicht möglich
End Sub

'Public Overridable Property ShowHelp() As Boolean
Public Property Get ShowHelp() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld Farbe die Hilfeschaltfläche
    ' angezeigt wird, oder legt diesen fest.
    ShowHelp = (m_Flags And CC_SHOWHELP)
End Property
Public Property Let ShowHelp(ByVal Value As Boolean)
    m_Flags = m_Flags Or CC_SHOWHELP
    If Not Value Then
        m_Flags = m_Flags Xor CC_SHOWHELP
    End If
End Property

'Public Overridable Property SolidColorOnly() As Boolean
Public Property Get SolidColorOnly() As Boolean
    ' Ruft einen Wert ab, der angibt, ob Benutzer im Dialogfeld ausschließlich
    ' Volltonfarben auswählen können, oder legt diesen fest.
     SolidColorOnly = (m_Flags And CC_SOLIDCOLOR)
End Property
Public Property Let SolidColorOnly(ByVal Value As Boolean)
    m_Flags = m_Flags Or CC_SOLIDCOLOR
    If Not Value Then
        m_Flags = m_Flags Xor CC_SOLIDCOLOR
    End If
End Property

'Public Overrides Function ToString() As String
Public Function ToStr() As String
    ToStr = "MBO.Windows.Forms.ColorDialog"
End Function

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult
    ReDim CancelColors(0 To MaxCustomColors - 1) As Long
Try: On Error GoTo Catch
    CancelColors() = m_CustomColors()
    Dim ahWndOwner As Long
    If IsMissing(Frm) Then
        ahWndOwner = GetActiveWindow
    Else
        ahWndOwner = Frm.hwnd
    End If
    Dim ChCl As ChooseColorW
    With ChCl
        .lStructSize = LenB(ChCl)
        'Debug.Print .lStructSize
        .hwndOwner = ahWndOwner
'#If VBA6 = 1 Or VBA7 = 1 Then
'        .hInstance = Application.HinstancePtr
'#Else
'        .hInstance = App.hInstance
'#End If
        .rgbResult = m_Color 'RGB(0, 255, 0) 'Farbe voreinstellen
        .lpCustColors = VarPtr(m_CustomColors(0)) 'Benutzerdefinierte Farben zuweisen
        .Flags = m_Flags
        .lCustData = 0&
        .lpfnHook = 0&
        .lpTemplateName = StrPtr(m_TemplateName) '"" '0&
    End With
    ShowDialog = ChooseColorW(ChCl)
    m_Color = ChCl.rgbResult
    If ShowDialog = 0 Then
        m_CustomColors() = CancelColors()
        CheckMessError
    End If
    Exit Function
Catch:
    ShowDialog = vbAbort
End Function

Private Sub CheckMessError()
    Dim hr As Long: hr = CommDlgExtendedError
    If hr = 0 Then Exit Sub
    Dim S As String
    Select Case hr
    Case CDERR_STRUCTSIZE:      S = "STRUCTSIZE"
    Case CDERR_INITIALIZATION:  S = "INITIALIZATION"
    Case CDERR_NOTEMPLATE:      S = "NOTEMPLATE"
    Case CDERR_NOHINSTANCE:     S = "NOHINSTANCE"
    Case CDERR_LOADSTRFAILURE:  S = "LOADSTRFAILURE"
    Case CDERR_FINDRESFAILURE:  S = "FINDRESFAILURE"
    Case CDERR_LOADRESFAILURE:  S = "LOADRESFAILURE"
    Case CDERR_LOCKRESFAILURE:  S = "LOCKRESFAILURE"
    Case CDERR_MEMALLOCFAILURE: S = "MEMALLOCFAILURE"
    Case CDERR_MEMLOCKFAILURE:  S = "MEMLOCKFAILURE"
    Case CDERR_NOHOOK:          S = "NOHOOK"
    Case CDERR_DIALOGFAILURE:   S = "DIALOGFAILURE"
    End Select
    If Len(S) Then
        MsgBox S
    End If
End Sub
