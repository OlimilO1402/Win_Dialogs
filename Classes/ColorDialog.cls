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
' Stellt ein Standarddialogfeld dar, in dem die verf�gbaren Farben angezeigt werden
' wie auch Steuerelemente, mit denen Benutzer benutzerdefinierte Farben definieren
' k�nnen.
Option Explicit
'Implements CommonDialog
Private Declare Function ChooseColorA Lib "comdlg32.dll" (pChooseColor As ChooseColor) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Type ChooseColor
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Private Const CC_RGBINIT              As Long = &H1   ' Vorgabe einer Standard-Farbe
Private Const CC_FULLOPEN             As Long = &H2   ' Vollauswahl aller Farben anzeigen
Private Const CC_PREVENTFULLOPEN      As Long = &H4   ' Deaktiviert den Button zum �ffnen der Dialogbox-Erweiterung
Private Const CC_SHOWHELP             As Long = &H8   ' Hilfe-Button anzeigen
Private Const CC_ENABLEHOOK           As Long = &H10  ' Nachrichten k�nnen "abgefangen" werden
Private Const CC_ENABLETEMPLATE       As Long = &H20  ' Dialogbox Template
Private Const CC_ENABLETEMPLATEHANDLE As Long = &H40  ' Benutzt Template, ignoriert aber den Template-Namen
Private Const CC_SOLIDCOLOR           As Long = &H80  ' nur Grundfarben ausw�hlbar
Private Const CC_ANYCOLOR             As Long = &H100 ' Anwender kann alle Farben w�hlen

Private mFlags As Long
Private Const MaxCustomColors = 16
Private mCustomColors(MaxCustomColors - 1) As Long 'mehr als 16 sind im Dialog gar nicht m�glich
Private mCstCclColors(MaxCustomColors - 1) As Long 'CancelColors
'wenn der Cancel-Button gedr�ckt wurde dann soll auch die alte Farbtabelle wiederhergestellt werden
Private mColor As Long

Private Sub Class_Initialize()
  '
  Call Reset
End Sub
'Public Sub New() 'keine Parameter, keine externe Constructor Funktion erforderlich

'Public Overridable Property AllowFullOpen() As Boolean
Public Property Get AllowFullOpen() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld benutzerdefinierte Farben definiert
' werden k�nnen, oder legt diesen fest.
   AllowFullOpen = Not (mFlags And CC_PREVENTFULLOPEN)
End Property
Public Property Let AllowFullOpen(BolVal As Boolean)
  mFlags = mFlags Or CC_PREVENTFULLOPEN
  If BolVal Then
    mFlags = mFlags Xor CC_PREVENTFULLOPEN
  End If
End Property
'Public Overridable Property FullOpen() As Boolean
Public Property Get FullOpen() As Boolean
' Ruft einen Wert ab, der angibt, ob die Steuerelemente f�r das Erstellen
' benutzerdefinierter Farben beim �ffnen des Dialogfelds angezeigt werden,
' oder legt diesen fest.
   FullOpen = Not (mFlags And CC_FULLOPEN)
End Property
Public Property Let FullOpen(BolVal As Boolean)
  mFlags = mFlags Or CC_FULLOPEN
  If Not BolVal Then
    mFlags = mFlags Xor CC_FULLOPEN
  End If
End Property

'Public Overridable Property AnyColor() As Boolean
Public Property Get AnyColor() As Boolean
' Ruft einen Wert ab, der angibt, ob im Dialogfeld bei den Grundfarben alle verf�gbaren
' Farben angezeigt werden, oder legt diesen fest.
   AnyColor = (mFlags And CC_ANYCOLOR)
End Property
Public Property Let AnyColor(BolVal As Boolean)
  mFlags = mFlags Or CC_ANYCOLOR
  If Not BolVal Then
    mFlags = mFlags Xor CC_ANYCOLOR
  End If
End Property

'Public Property Color() As System.Drawing.Color
Public Property Get Color() As Long 'System.Drawing.Color
' Ruft die von den Benutzern ausgew�hlte Farbe ab oder legt diese fest.
  Color = mColor
End Property
Public Property Let Color(ColVal As Long) 'System.Drawing.Color
  mColor = ColVal
End Property

'Public Property CustomColors() As Integer()
Public Property Get CustomColors(nIndex As Long) As Long 'Integer()
' Ruft den im Dialogfeld angezeigten Satz benutzerdefinierter Farben ab
' oder legt diesen fest.
  CustomColors = mCustomColors(nIndex)
End Property
Public Property Let CustomColors(nIndex As Long, ColVal As Long) 'Integer()
  mCustomColors(nIndex) = ColVal
End Property

'Public Overrides Sub Reset()
Public Sub Reset()
' Setzt alle Optionen auf die Standardwerte, die zuletzt ausgew�hlte Farbe auf
' schwarz und die benutzerdefinierten Farben auf die Standardwerte zur�ck.
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
' Ruft einen Wert ab, der angibt, ob im Dialogfeld Farbe die Hilfeschaltfl�che
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
  ' Ruft einen Wert ab, der angibt, ob Benutzer im Dialogfeld ausschlie�lich
  ' Volltonfarben ausw�hlen k�nnen, oder legt diesen fest.
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
Dim ChCl As ChooseColor
Dim hhwndOwner As Long
Dim CancelColors(MaxCustomColors - 1) As Long, i As Long
TryE: On Error GoTo CatchE
'die Custom Farbtabelle sichern 'falls Abbrechen geklickt wurde
  For i = 0 To MaxCustomColors - 1
    CancelColors(i) = mCustomColors(i)
  Next
  If IsMissing(Frm) Then
    hhwndOwner = GetActiveWindow
  Else
    hhwndOwner = Frm.hwnd
  End If
  With ChCl
    .lStructSize = Len(ChCl)
    .hwndOwner = hhwndOwner
    .hInstance = App.hInstance
    .rgbResult = mColor 'RGB(0, 255, 0) 'Farbe voreinstellen
    .lpCustColors = VarPtr(mCustomColors(0)) 'Benutzerdefinierte Farben zuweisen
    .flags = mFlags
    .lCustData = 0&
    .lpfnHook = 0&
    '.lpTemplateName = "" '0&
  End With
  ShowDialog = ChooseColorA(ChCl)
  mColor = ChCl.rgbResult
  If ShowDialog = vbCancel Then
    For i = 0 To MaxCustomColors - 1
      mCustomColors(i) = CancelColors(i)
    Next
  End If
  Exit Function
CatchE:
  ShowDialog = vbAbort
End Function

