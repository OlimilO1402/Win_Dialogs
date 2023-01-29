VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIFileDialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 12/2018

' Komplett

Option Explicit

Private Const ERR_BASE As Long = 40670
Private Const CLSID_FileOpenDialog As String = "{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}"
Private Const CLSID_FileSaveDialog As String = "{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}"
Private Const IID_IFileDialog2 As String = "{61744fc7-85b5-4791-a9b0-272276309b13}"

Private Enum E_vtbOffsets
    '/*** IUnknown methods ***/
    vtb_QueryInterface
    vtb_AddRef
    vtb_Release
    '/*** IFileDialog methods ***/
    vtb_Show
    vtb_SetFileTypes
    vtb_SetFileTypeIndex
    vtb_GetFileTypeIndex
    vtb_Advise
    vtb_Unadvise
    vtb_SetOptions
    vtb_GetOptions
    vtb_SetDefaultFolder
    vtb_SetFolder
    vtb_GetFolder
    vtb_GetCurrentSelection
    vtb_SetFileName
    vtb_GetFileName
    vtb_SetTitle
    vtb_SetOkButtonLabel
    vtb_SetFileNameLabel
    vtb_GetResult
    vtb_AddPlace
    vtb_SetDefaultExtension
    vtb_Close
    vtb_SetClientGuid
    vtb_ClearClientData
    vtb_SetFilter
    '/*** IFileDialog2 methods ***/
    vtb_SetCancelButtonLabel
    vtb_SetNavigationRoot
End Enum

Private m_initialized As Boolean
Private m_Interface As Interface_Data

' ----==== Interface IUnknown Func ====----
Public Function QueryInterface(ByVal riid As String) As Long

    Dim pRet As Long
    Dim triid(3) As Long
    Dim psriid As Long
    Dim ptriid As Long
    Dim ppvObject As Long
    Dim pppvObject As Long

    psriid = StrPtr(riid)
    ptriid = VarPtr(triid(0))

    Call IIDFromString(psriid, ptriid)

    pppvObject = VarPtr(ppvObject)

    If x_OleInvoke(vtb_QueryInterface, ptriid, pppvObject) Then

        pRet = ppvObject

    End If

    QueryInterface = pRet

End Function

' ----==== Interface IFileDialog Func ====----
Public Function Show(ByVal hwndOwner As Long) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_Show, hwndOwner) Then

        bRet = True

    End If

    Show = bRet

End Function

Friend Function SetFileTypes(ByRef rgFilterSpec() As COMDLG_FILTERSPEC) As Boolean

    Dim bRet As Boolean
    Dim cFileTypes As Long
    Dim prgFilterSpec As Long

    bRet = False
    cFileTypes = UBound(rgFilterSpec) + 1
    prgFilterSpec = VarPtr(rgFilterSpec(0))

    If x_OleInvoke(vtb_SetFileTypes, cFileTypes, prgFilterSpec) Then

        bRet = True

    End If

    SetFileTypes = bRet

End Function

Public Function SetFileTypeIndex(ByVal iFileType As Long) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_SetFileTypeIndex, iFileType) Then

        bRet = True

    End If

    SetFileTypeIndex = bRet

End Function

Public Function GetFileTypeIndex() As Long

    Dim lRet As Long
    Dim iFileType As Long
    Dim piFileType As Long

    piFileType = VarPtr(iFileType)

    If x_OleInvoke(vtb_GetFileTypeIndex, piFileType) Then

        lRet = iFileType

    End If

    GetFileTypeIndex = lRet

End Function

Public Function Advise(ByVal pfde As Long) As Long

    Dim pRet As Long
    Dim dwCookie As Long
    Dim pdwCookie As Long

    pdwCookie = VarPtr(dwCookie)

    If x_OleInvoke(vtb_Advise, pfde, pdwCookie) Then

        pRet = dwCookie

    End If

    Advise = pRet

End Function

Public Function Unadvise(ByVal dwCookie As Long) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_Unadvise, dwCookie) Then

        bRet = True

    End If

    Unadvise = bRet

End Function

Public Function SetOptions(ByVal fos As FILEOPENDIALOGOPTIONS) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_SetOptions, fos) Then

        bRet = True

    End If

    SetOptions = bRet

End Function

Public Function GetOptions() As FILEOPENDIALOGOPTIONS

    Dim lRet As Long
    Dim fos As Long
    Dim pfos As Long

    pfos = VarPtr(fos)

    If x_OleInvoke(vtb_GetOptions, pfos) Then

        lRet = fos

    End If

    GetOptions = lRet

End Function

Public Function SetDefaultFolder(ByVal psi As Long) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_SetDefaultFolder, psi) Then

        bRet = True

    End If

    SetDefaultFolder = bRet

End Function

Public Function SetFolder(ByVal psi As Long) As Boolean

    Dim bRet As Long

    bRet = False

    If x_OleInvoke(vtb_SetFolder, psi) Then

        bRet = True

    End If

    SetFolder = bRet

End Function

Public Function GetFolder() As Long

    Dim pRet As Long
    Dim psi As Long
    Dim ppsi As Long

    ppsi = VarPtr(psi)

    If x_OleInvoke(vtb_GetFolder, ppsi) Then

        pRet = psi

    End If

    GetFolder = pRet

End Function

Public Function GetCurrentSelection() As Long

    Dim pRet As Long
    Dim psi As Long
    Dim ppsi As Long

    ppsi = VarPtr(psi)

    If x_OleInvoke(vtb_GetCurrentSelection, ppsi) Then

        pRet = psi

    End If

    GetCurrentSelection = pRet

End Function

Public Function SetFileName(ByVal szName As String) As Boolean

    Dim bRet As Boolean
    Dim pszName As Long

    bRet = False
    pszName = StrPtr(szName)

    If x_OleInvoke(vtb_SetFileName, pszName) Then

        bRet = True

    End If

    SetFileName = bRet

End Function

Public Function GetFileName() As String

    Dim sRet As String
    Dim szName As Long
    Dim pszName As Long

    pszName = VarPtr(szName)

    If x_OleInvoke(vtb_GetFileName, pszName) Then

        sRet = GetStringFromPointer(szName)

    End If

    GetFileName = sRet

End Function

Public Function SetTitle(ByVal szTitle As String) As Boolean

    Dim bRet As Boolean
    Dim pszTitle As Long

    bRet = False
    pszTitle = StrPtr(szTitle)

    If x_OleInvoke(vtb_SetTitle, pszTitle) Then

        bRet = True

    End If

    SetTitle = bRet

End Function

Public Function SetOkButtonLabel(ByVal szText As String) As Boolean

    Dim bRet As Boolean
    Dim pszText As Long

    bRet = False
    pszText = StrPtr(szText)

    If x_OleInvoke(vtb_SetOkButtonLabel, pszText) Then

        bRet = True

    End If

    SetOkButtonLabel = bRet

End Function

Public Function SetFileNameLabel(ByVal szLabel As String) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False
    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_SetFileNameLabel, pszLabel) Then

        bRet = True

    End If

    SetFileNameLabel = bRet

End Function

Public Function GetResult() As Long

    Dim pRet As Long
    Dim ppsi As Long
    Dim pppsi As Long

    pppsi = VarPtr(ppsi)

    If x_OleInvoke(vtb_GetResult, pppsi) Then

        pRet = ppsi

    End If

    GetResult = pRet

End Function

Public Function AddPlace(ByVal psi As Long, ByVal efdap As FDAP) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_AddPlace, psi, efdap) Then

        bRet = True

    End If

    AddPlace = bRet

End Function

Public Function SetDefaultExtension(ByVal szDefaultExtension As String) As Boolean

    Dim bRet As Boolean
    Dim pszDefaultExtension As Long

    bRet = False
    pszDefaultExtension = StrPtr(szDefaultExtension)

    If x_OleInvoke(vtb_SetDefaultExtension, pszDefaultExtension) Then

        bRet = True

    End If

    SetDefaultExtension = bRet

End Function

Public Function CloseDlg() As Long

    Dim lRet As Long
    Dim hr As Long
    Dim phr As Long

    phr = VarPtr(hr)

    If x_OleInvoke(vtb_Close, phr) Then

        lRet = hr

    End If

    CloseDlg = lRet

End Function

Friend Function SetClientGuid(ByRef tguid As GUID) As Boolean

    Dim bRet As Boolean
    Dim ptguid As Long

    bRet = False
    ptguid = VarPtr(tguid)

    If x_OleInvoke(vtb_SetClientGuid, ptguid) Then

        bRet = True

    End If

    SetClientGuid = bRet

End Function

Public Function ClearClientData() As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_ClearClientData) Then

        bRet = True

    End If

    ClearClientData = bRet

End Function

' [Deprecated. SetFilter is no longer available for use as of Windows 7.]
Public Function SetFilter(ByVal psi As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetFilter, psi) Then

        bRet = True

    End If

    SetFilter = bRet

End Function

' ----==== Interface IFileDialog2 Func ====----
Public Function SetCancelButtonLabel(ByVal szLabel As String) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False
    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_SetCancelButtonLabel, pszLabel) Then

        bRet = True

    End If

    SetCancelButtonLabel = bRet

End Function

Public Function SetNavigationRoot(ByVal psi As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetNavigationRoot, psi) Then

        bRet = True

    End If

    SetNavigationRoot = bRet

End Function

' ----==== Interface Helper Func ====----
Public Sub InitAs(Optional ByVal eDialog As DialogType = FileOpenDialog)

    m_Interface.owner = ObjPtr(Me)
    m_Interface.RaiseErrors = False

    Select Case eDialog

    Case DialogType.FileOpenDialog

        m_initialized = InitInterface(m_Interface, CLSID_FileOpenDialog, _
            IID_IFileDialog2)

    Case DialogType.FileSaveDialog

        m_initialized = InitInterface(m_Interface, CLSID_FileSaveDialog, _
            IID_IFileDialog2)

    End Select

End Sub

Private Sub Class_Terminate()

    ReleaseInterface m_Interface

End Sub

Public Sub RaiseError(Optional ByVal Raise As Boolean = True)

    m_Interface.RaiseErrors = Raise

End Sub

Private Function x_OleInvoke(ByVal vtb As E_vtbOffsets, ParamArray var()) As Boolean

    Dim ret As Long

    If Not m_initialized Then

        InterfaceError m_Interface, ecd_OleInvoke

    Else

        x_OleInvoke = oleInvoke(m_Interface, vtb, ret, True, var)

    End If

End Function

Public Sub x_RaiseError()

    With m_Interface

        If .ecd Then Err.Raise ERR_BASE + .ecd, App.EXEName & _
            ".clsIFileDialog2", .etx

        .ecd = 0
        .etx = ""

    End With

End Sub



