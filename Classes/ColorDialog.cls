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
Private Declare Function ChooseColorW Lib "comdlg32.dll" (pChooseColor As ChooseColor) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Type ChooseColor
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type
Private Const CC_RGBINIT              As Long = &H1   ' Vorgabe einer Standard-Farbe
Private Const CC_FULLOPEN             As Long = &H2   ' Vollauswahl aller Farben anzeigen
Private Const CC_PREVENTFULLOPEN      As Long = &H4   ' Deaktiviert den Button zum Öffnen der Dialogbox-Erweiterung
Private Const CC_SHOWHELP             As Long = &H8   ' Hilfe-Button anzeigen
Private Const CC_ENABLEHOOK           As Long = &H10  ' Nachrichten können "abgefangen" werden
Private Const CC_ENABLETEMPLATE       As Long = &H20  ' Dialogbox Template
Private Const CC_ENABLETEMPLATEHANDLE As Long = &H40  ' Benutzt Template, ignoriert aber den Template-Namen
Private Const CC_SOLIDCOLOR           As Long = &H80  ' nur Grundfarben auswählbar
Private Const CC_ANYCOLOR             As Long = &H100 ' Anwender kann alle Farben wählen

Private mFlags As Long
Private Const MaxCustomColors As Long = 16
Private mCustomColors(MaxCustomColors - 1) As Long 'mehr als 16 sind im Dialog gar nicht möglich
Private mCstCclColors(MaxCustomColors - 1) As Long 'CancelColors
'wenn der Cancel-Button gedrückt wurde dann soll auch die alte Farbtabelle wiederhergestellt werden
Private mColor As Long

Private Sub Class_Initialize()
    '
    Call Reset
End Sub
'Public Sub New() 'keine Parameter, keine externe Constructor Funktion erforderlich

'Public Overridable Property AllowFullOpen() As Boolean
Public Property Get AllowFullOpen() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld benutzerdefinierte Farben definiert
    ' werden können, oder legt diesen fest.
    AllowFullOpen = Not (mFlags And CC_PREVENTFULLOPEN)
End Property
Public Property Let AllowFullOpen(ByVal Value As Boolean)
    mFlags = mFlags Or CC_PREVENTFULLOPEN
    If Value Then
        mFlags = mFlags Xor CC_PREVENTFULLOPEN
    End If
End Property
'Public Overridable Property FullOpen() As Boolean
Public Property Get FullOpen() As Boolean
    ' Ruft einen Wert ab, der angibt, ob die Steuerelemente für das Erstellen
    ' benutzerdefinierter Farben beim Öffnen des Dialogfelds angezeigt werden,
    ' oder legt diesen fest.
    FullOpen = Not (mFlags And CC_FULLOPEN)
End Property
Public Property Let FullOpen(ByVal Value As Boolean)
    mFlags = mFlags Or CC_FULLOPEN
    If Not Value Then
        mFlags = mFlags Xor CC_FULLOPEN
    End If
End Property

'Public Overridable Property AnyColor() As Boolean
Public Property Get AnyColor() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld bei den Grundfarben alle verfügbaren
    ' Farben angezeigt werden, oder legt diesen fest.
    AnyColor = (mFlags And CC_ANYCOLOR)
End Property
Public Property Let AnyColor(ByVal Value As Boolean)
    mFlags = mFlags Or CC_ANYCOLOR
    If Not Value Then
        mFlags = mFlags Xor CC_ANYCOLOR
    End If
End Property

'Public Property Color() As System.Drawing.Color
Public Property Get Color() As Long 'System.Drawing.Color
    ' Ruft die von den Benutzern ausgewählte Farbe ab oder legt diese fest.
    Color = mColor
End Property
Public Property Let Color(ByVal Value As Long)  'System.Drawing.Color
    mColor = Value
End Property

'Public Property CustomColors() As Integer()
Public Property Get CustomColors(ByVal Index As Long) As Long 'Integer()
    ' Ruft den im Dialogfeld angezeigten Satz benutzerdefinierter Farben ab
    ' oder legt diesen fest.
    If Index > 16 Then Exit Property
    CustomColors = mCustomColors(Index)
End Property
Public Property Let CustomColors(ByVal Index As Long, ByVal Value As Long) 'Integer()
    If Index > 16 Then Exit Property
    mCustomColors(Index) = Value
End Property

'Public Overrides Sub Reset()
Public Sub Reset()
    ' Setzt alle Optionen auf die Standardwerte, die zuletzt ausgewählte Farbe auf
    ' schwarz und die benutzerdefinierten Farben auf die Standardwerte zurück.
    ' welche Standardwerte der UD Farben?
    mFlags = mFlags Or CC_RGBINIT
    mFlags = mFlags Or CC_ANYCOLOR
    mFlags = mFlags Or CC_FULLOPEN
    mFlags = mFlags Or CC_PREVENTFULLOPEN
    mColor = vbBlack '&H0& 'Schwarz '
    'ReDim mCustomColors(16)
End Sub

'Public Overridable Property ShowHelp() As Boolean
Public Property Get ShowHelp() As Boolean
    ' Ruft einen Wert ab, der angibt, ob im Dialogfeld Farbe die Hilfeschaltfläche
    ' angezeigt wird, oder legt diesen fest.
    ShowHelp = (mFlags And CC_SHOWHELP)
End Property
Public Property Let ShowHelp(BolVal As Boolean)
    mFlags = mFlags Or CC_SHOWHELP
    If Not BolVal Then
        mFlags = mFlags Xor CC_SHOWHELP
    End If
End Property

'Public Overridable Property SolidColorOnly() As Boolean
Public Property Get SolidColorOnly() As Boolean
    ' Ruft einen Wert ab, der angibt, ob Benutzer im Dialogfeld ausschließlich
    ' Volltonfarben auswählen können, oder legt diesen fest.
     SolidColorOnly = (mFlags And CC_SOLIDCOLOR)
End Property
Public Property Let SolidColorOnly(BolVal As Boolean)
    mFlags = mFlags Or CC_SOLIDCOLOR
    If Not BolVal Then
        mFlags = mFlags Xor CC_SOLIDCOLOR
    End If
End Property

'Public Overrides Function ToString() As String
Public Function ToStr() As String
    ToStr = "MBO.Windows.Forms.ColorDialog"
End Function

Public Function ShowDialog(Optional Frm As Variant) As VbMsgBoxResult
    ReDim CancelColors(MaxCustomColors - 1) As Long
    Dim i As Long
Try: On Error GoTo Catch
    'die Custom Farbtabelle sichern 'falls Abbrechen geklickt wurde
    'For i = 0 To MaxCustomColors - 1
    '    CancelColors(i) = mCustomColors(i)
    'Next
    CancelColors() = mCustomColors()
    Dim hhwndOwner As Long
    If IsMissing(Frm) Then
        hhwndOwner = GetActiveWindow
    Else
        hhwndOwner = Frm.hwnd
    End If
    Dim ChCl As ChooseColor
    With ChCl
        .lStructSize = LenB(ChCl)
        .hwndOwner = hhwndOwner
        .hInstance = App.hInstance
        .rgbResult = mColor 'RGB(0, 255, 0) 'Farbe voreinstellen
        .lpCustColors = VarPtr(mCustomColors(0)) 'Benutzerdefinierte Farben zuweisen
        .flags = mFlags
        .lCustData = 0&
        .lpfnHook = 0&
        '.lpTemplateName = "" '0&
    End With
    ShowDialog = ChooseColorW(ChCl)
    mColor = ChCl.rgbResult
    If ShowDialog = vbCancel Then
        For i = 0 To MaxCustomColors - 1
            mCustomColors(i) = CancelColors(i)
        Next
    End If
    Exit Function
Catch:
    ShowDialog = vbAbort
End Function

