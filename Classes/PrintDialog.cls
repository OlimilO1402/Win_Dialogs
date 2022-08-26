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

'Oh wowo if you deal with PrintDialog you come across a whole new universe of printer settings and print settings, paper size and stuff

'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms646942(v=vs.85)
'https://docs.microsoft.com/de-de/windows/win32/api/commdlg/ns-commdlg-printdlgexw

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type PRINTDLG
    lStructSize   As Long
    hwndOwner     As LongPtr
    hdevmode      As LongPtr
    hDevNames     As LongPtr
    hDC           As LongPtr
    Flags         As Long
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

Private Type PRINTDLGEXW
    lStructSize      As Long
    hwndOwner        As LongPtr '  HWND
    hdevmode         As LongPtr '  HGLOBAL
    hDevNames        As LongPtr '  HGLOBAL
    hDC              As LongPtr '  HDC
    Flags            As Long    '  DWORD
    Flags2           As Long    '  DWORD
    ExclusionFlags   As Long    '  DWORD
    nPageRanges      As Long    '  DWORD
    nMaxPageRanges   As Long    '  DWORD
    lpPageRanges     As LongPtr '  LPPRINTPAGERANGE
    nMinPage         As Long    '  DWORD
    nMaxPage         As Long    '  DWORD
    nCopies          As Long    '  DWORD
    hInstance        As LongPtr '  HINSTANCE
    lpPrintTemplateName As LongPtr '  LPCWSTR
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
    pszHeaderTitle As LongPtr 'LPCWSTR
    pszHeaderSubTitle As LongPtr '  LPCWSTR
    hActCtx As LongPtr '  HANDLE
'  union {
'    HBITMAP hbmHeader;
'    LPCWSTR pszbmHeader;
'  } DUMMYUNIONNAME3;
End Type
Private Declare Function PrintDlgW Lib "comdlg32" (ByRef pPrintdlg As PRINTDLG) As Long
Private Declare Function PRINTDLGEXW Lib "comdlg32" Alias "PrintDlgExW" (ByRef TLPPRINTDLGEXW As PRINTDLGEXW) As Long
Private Declare Function PrinterMessageBoxW Lib "winspool.drv" ( _
    ByVal hPrinter As Long, _
    ByVal error As Long, _
    ByVal hwnd As Long, _
    ByVal pText As String, _
    ByVal pCaption As String, _
    ByVal dwType As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" ( _
    ByVal hwnd As Long, _
    ByVal hPrinter As Long) As Long


'Properties of Windows.Forms.PrintDialog
Public Property Get AllowCurrentPage() As Boolean
    '
End Property
Public Property Get AllowPrintToFile() As Boolean
    '
End Property
Public Property Get AllowSelection() As Boolean
    '
End Property
Public Property Get AllowSomePages() As Boolean
    '
End Property
Public Property Get UseEXDialog() As Boolean
    '
End Property
Public Property Get ShowNetwork() As Boolean
    '
End Property
'Public Property Get Container() As IContainer
'    '
'End Property
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
    '
End Property
'Public Property Get PrinterSettings_DefaultPageSettings_Bounds() As Rectangle
'    '
'End Property
Public Property Get PrinterSettings_DefaultPageSettings_Clone() As Object
    '
End Property
Public Property Get PrinterSettings_DefaultPageSettings_Color() As Boolean
    '
End Property
'Public Property Get PrinterSettings_DefaultPageSettings_CopyToHdevmode(hdevmode As IntPtr)
'    '
'End Property
'Public Property Get PrinterSettings_Duplex() As Duplex 'Enum
'    '
'End Property
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


Public Property Get PrinterSettings_FromPage() As Integer
    '
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
Public Property Get PrinterSettings_MaximumPage() As Long
    '
End Property
Public Property Get PrinterSettings_MinimumPage() As Long
    '
End Property
'Public Property Get PrinterSettings_PaperSizes() As PaperSizeCollection
'    '
'End Property
'Public Property Get PrinterSettings_PaperSources() As PaperSourceCollection
'    '
'End Property
Public Property Get PrinterSettings_PrinterName() As String
    '
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
Public Property Get PrinterSettings_PrintToFile() As Boolean
    '
End Property
Public Property Get PrinterSettings_SupportsColor() As Boolean
    '
End Property
Public Property Get PrinterSettings_ToPage() As Long
    '
End Property
Public Property Get PrintToFile() As Boolean
    '
End Property
Public Sub Reset()
    '
End Sub
Public Function ShowDialog() As VbMsgBoxResult
    '
End Function
Public Property Get ShowHelp() As Boolean
    '
End Property
'Public Property Get ShowNetwork() As Boolean
'    '
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
'Public Property Get UseEXDialog() As Boolean
'    '
'End Property
'
