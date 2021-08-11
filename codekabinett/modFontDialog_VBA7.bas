Attribute VB_Name = "modFontDialog_VBA7"
Option Compare Database
Option Explicit

' Original Code by Terry Kreft
' Modified by Stephen Lebans (http://lebans.com/)
'
' This code was revised on 2019-01-26 in a hurry
' by Philipp Stiefel (http://codekabinett.com) to
' make it run in x64 VBA-Applications.
' This is just the code that "worked for me" with
' a few superficial tests. I did not conduct a
' serious code review and do not take any
' responsibility for any defects in the code.


'************  Code Start  ***********
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const LF_FACESIZE = 32

Private Const FW_BOLD = 700

Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_OEMTEXT = 7
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS

Public Const LOGPIXELSY = 90

Public Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As LongPtr
  hdc As LongPtr
  lpLogFont As LongPtr
  iPointSize As Long
  Flags As Long
  rgbColors As Long
  lCustData As LongPtr
  lpfnHook As LongPtr
  lpTemplateName As String
  hInstance As LongPtr
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
(pChoosefont As FONTSTRUC) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr


Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
Dim lngTemp As Long
  On Error GoTo MulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
MulDiv_end:
  MulDiv = lngTemp
  Exit Function
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err
End Function
Private Function ByteToString(aBytes() As Byte) As String
  Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
  dwBytePoint = LBound(aBytes)
  While dwBytePoint <= UBound(aBytes)
    dwByteVal = aBytes(dwBytePoint)
    If dwByteVal = 0 Then
      ByteToString = szOut
      Exit Function
    Else
      szOut = szOut & Chr$(dwByteVal)
    End If
    dwBytePoint = dwBytePoint + 1
  Wend
  ByteToString = szOut
End Function

Private Sub StringToByte(InString As String, ByteArray() As Byte)
Dim intLbound As Integer
  Dim intUbound As Integer
  Dim intLen As Integer
  Dim intX As Integer
  intLbound = LBound(ByteArray)
  intUbound = UBound(ByteArray)
  intLen = Len(InString)
  If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
For intX = 1 To intLen
ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
Next
End Sub


Public Function DialogFont(ByRef f As FormFontInfo) As Boolean
  Dim LF As LOGFONT, FS As FONTSTRUC
  Dim lLogFontAddress As LongPtr, lMemHandle As LongPtr

  LF.lfWeight = f.Weight
  LF.lfItalic = f.Italic * -1
  LF.lfUnderline = f.UnderLine * -1
  LF.lfHeight = -MulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
  Call StringToByte(f.Name, LF.lfFaceName())
  FS.rgbColors = f.Color
  FS.lStructSize = LenB(FS)

    ' To be modal must be valid Hwnd
    FS.hwnd = Application.hWndAccessApp
    
  lMemHandle = GlobalAlloc(GHND, LenB(LF))
  If lMemHandle = 0 Then
    DialogFont = False
    Exit Function
  End If

  lLogFontAddress = GlobalLock(lMemHandle)
  If lLogFontAddress = 0 Then
    DialogFont = False
    Exit Function
  End If

  CopyMemory ByVal lLogFontAddress, LF, LenB(LF)
  FS.lpLogFont = lLogFontAddress
  FS.Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
  Dim apiRetVal As Long
  apiRetVal = ChooseFont(FS)
  If apiRetVal = 1 Then
    CopyMemory LF, ByVal lLogFontAddress, LenB(LF)
    f.Weight = LF.lfWeight
    f.Italic = CBool(LF.lfItalic)
    f.UnderLine = CBool(LF.lfUnderline)
    f.Name = ByteToString(LF.lfFaceName())
    f.Height = CLng(FS.iPointSize / 10)
    f.Color = FS.rgbColors
    
    DialogFont = True
  Else
    DialogFont = False
  End If
End Function

Function test_DialogFont(ctl As Control) As Boolean
    Dim f As FormFontInfo
    With f
      .Color = 0
      .Height = 12
      .Weight = 700
      .Italic = False
      .UnderLine = False
      .Name = "Arial"
    End With
    Call DialogFont(f)
    With f
        Debug.Print "Font Name: "; .Name
        Debug.Print "Font Size: "; .Height
        Debug.Print "Font Weight: "; .Weight
        Debug.Print "Font Italics: "; .Italic
        Debug.Print "Font Underline: "; .UnderLine
        Debug.Print "Font COlor: "; .Color
        
        ctl.FontName = .Name
        ctl.FontSize = .Height
        ctl.FontWeight = .Weight
        ctl.FontItalic = .Italic
        ctl.FontUnderline = .UnderLine
        ctl = .Name & " - Size:" & .Height
    End With
    test_DialogFont = True
End Function
'************  Code End  ***********

