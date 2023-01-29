Attribute VB_Name = "modFunc"
Option Explicit

' ----==== Enums ====----
Public Enum DialogType
    FileOpenDialog = &H0
    FileSaveDialog = &H1
End Enum

Public Enum CDCONTROLSTATEF
    CDCS_INACTIVE = &H0
    CDCS_ENABLED = &H1
    CDCS_VISIBLE = &H2
    CDCS_ENABLEDVISIBLE = &H3
End Enum

Public Enum FDAP
    FDAP_BOTTOM = &H0
    FDAP_TOP = &H1
End Enum

Public Enum FILEOPENDIALOGOPTIONS
    FOS_OVERWRITEPROMPT = &H2
    FOS_STRICTFILETYPES = &H4
    FOS_NOCHANGEDIR = &H8
    FOS_PICKFOLDERS = &H20
    FOS_FORCEFILESYSTEM = &H40
    FOS_ALLNONSTORAGEITEMS = &H80
    FOS_NOVALIDATE = &H100
    FOS_ALLOWMULTISELECT = &H200
    FOS_PATHMUSTEXIST = &H800
    FOS_FILEMUSTEXIST = &H1000
    FOS_CREATEPROMPT = &H2000
    FOS_SHAREAWARE = &H4000
    FOS_NOREADONLYRETURN = &H8000
    FOS_NOTESTFILECREATE = &H10000
    FOS_HIDEMRUPLACES = &H20000
    FOS_HIDEPINNEDPLACES = &H40000
    FOS_NODEREFERENCELINKS = &H100000
    FOS_OKBUTTONNEEDSINTERACTION = &H200000
    FOS_DONTADDTORECENT = &H2000000
    FOS_FORCESHOWHIDDEN = &H10000000
    FOS_DEFAULTNOMINIMODE = &H20000000
    FOS_FORCEPREVIEWPANEON = &H40000000
    FOS_SUPPORTSTREAMABLEITEMS = &H80000000
End Enum

Public Enum SIGDN
    SIGDN_NORMALDISPLAY = &H0
    SIGDN_PARENTRELATIVEPARSING = &H80018001
    SIGDN_DESKTOPABSOLUTEPARSING = &H80028000
    SIGDN_PARENTRELATIVEEDITING = &H80031001
    SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000
    SIGDN_FILESYSPATH = &H80058000
    SIGDN_URL = &H80068000
    SIGDN_PARENTRELATIVEFORADDRESSBAR = &H8007C001
    SIGDN_PARENTRELATIVE = &H80080001
    SIGDN_PARENTRELATIVEFORUI = &H80094001
End Enum

Public Enum SIATTRIBFLAGS
    SIATTRIBFLAGS_AND = 1
    SIATTRIBFLAGS_APPCOMPAT = 3
    SIATTRIBFLAGS_OR = 2
End Enum

Public Enum GETPROPERTYSTOREFLAGS
    GPS_DEFAULT = &H0
    GPS_HANDLERPROPERTIESONLY = &H1
    GPS_READWRITE = &H2
    GPS_TEMPORARY = &H4
    GPS_FASTPROPERTIESONLY = &H8
    GPS_OPENSLOWITEM = &H10
    GPS_DELAYCREATION = &H20
    GPS_BESTEFFORT = &H40
    GPS_NO_OPLOCK = &H80
    GPS_MASK_VALID = &HFF
End Enum

Public Enum SFGAOF
    SFGAO_CANCOPY = &H1
    SFGAO_CANMOVE = &H2
    SFGAO_CANLINK = &H4
    SFGAO_STORAGE = &H8
    SFGAO_CANRENAME = &H10
    SFGAO_CANDELETE = &H20
    SFGAO_HASPROPSHEET = &H40
    SFGAO_DROPTARGET = &H100
    SFGAO_CAPABILITYMASK = &H177
    SFGAO_ENCRYPTED = &H2000
    SFGAO_ISSLOW = &H4000
    SFGAO_GHOSTED = &H8000
    SFGAO_LINK = &H10000
    SFGAO_SHARE = &H20000
    SFGAO_READONLY = &H40000
    SFGAO_HIDDEN = &H80000
    SFGAO_DISPLAYATTRMASK = &HFC000
    SFGAO_FILESYSANCESTOR = &H10000000
    SFGAO_FOLDER = &H20000000
    SFGAO_FILESYSTEM = &H40000000
    SFGAO_HASSUBFOLDER = &H80000000
    SFGAO_CONTENTSMASK = &H80000000
    SFGAO_VALIDATE = &H1000000
    SFGAO_REMOVABLE = &H2000000
    SFGAO_COMPRESSED = &H4000000
    SFGAO_BROWSABLE = &H8000000
    SFGAO_NONENUMERATED = &H100000
    SFGAO_NEWCONTENT = &H200000
    SFGAO_CANMONIKER = &H400000
    SFGAO_HASSTORAGE = &H400000
    SFGAO_STREAM = &H400000
    SFGAO_STORAGEANCESTOR = &H800000
    SFGAO_STORAGECAPMASK = &H70C50008
End Enum

Public Enum SICHINTF
    SICHINT_DISPLAY = &H0
    SICHINT_ALLFIELDS = &H80000000
    SICHINT_CANONICAL = &H10000000
    SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL = &H20000000
End Enum

' ----==== Types ====----
Public Type COMDLG_FILTERSPEC
    pszName As String
    pszSpec As String
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type PROPERTYKEY
    fmtid As GUID
    pid As Long
End Type

' ----==== Ole32 API-Deklarationen ====----
Public Declare Sub CoTaskMemFree Lib "ole32.dll" ( _
                   ByVal hMem As Long)

Public Declare Function IIDFromString Lib "ole32" ( _
                         ByVal lpsz As Long, _
                         ByVal lpiid As Long) As Long

Public Declare Function StringFromCLSID Lib "ole32.dll" ( _
                        ByRef pCLSID As GUID, _
                        ByRef lpszProgID As Long) As Long

' ----==== Kernel32 API-Deklarationen ====----
Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
                   ByRef hpvDest As Any, _
                   ByRef hpvSource As Any, _
                   ByVal cbCopy As Long)

Private Declare Function lstrlenW Lib "kernel32" ( _
                         ByVal lpString As Long) As Long

' ----==== StrPtr to String ====----
Public Function GetStringFromPointer(ByVal lpStrPointer As Long) As String

    Dim lLen As Long
    Dim bBuffer() As Byte

    lLen = lstrlenW(lpStrPointer) * 2 - 1

    If lLen > 0 Then

        ReDim bBuffer(lLen)

        Call RtlMoveMemory(bBuffer(0), ByVal lpStrPointer, lLen)

        Call CoTaskMemFree(lpStrPointer)

        GetStringFromPointer = bBuffer

    End If

End Function

' ----==== GUID to String ====----
Public Function Guid2String(ByRef tguid As GUID) As String

    Dim lGuid As Long
    
    Call StringFromCLSID(tguid, lGuid)

    Guid2String = GetStringFromPointer(lGuid)

End Function

