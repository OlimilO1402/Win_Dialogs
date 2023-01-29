Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Public Class PickFolderDialog
    Implements IDisposable
    Implements IFileDialogEvents

#Region "New"
    Public Sub New()
        'Dim eFILEOPENDIALOGOPTIONS As FILEOPENDIALOGOPTIONS
        m_FileOpenDialog = Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid(CLSID_FileOpenDialog)))
        If m_FileOpenDialog IsNot Nothing Then
            'If CType(m_FileOpenDialog, IFileDialog2).GetOptions(eFILEOPENDIALOGOPTIONS) = S_OK Then
            '    eFILEOPENDIALOGOPTIONS = eFILEOPENDIALOGOPTIONS Or FILEOPENDIALOGOPTIONS.FOS_PICKFOLDERS
            '    If CType(m_FileOpenDialog, IFileDialog2).SetOptions(eFILEOPENDIALOGOPTIONS) = S_OK Then
            If CType(m_FileOpenDialog, IFileDialog2).Advise(Me, m_Cookie) = S_OK Then
            End If
            '    End If
            'End If
        End If
    End Sub
#End Region

#Region "API"
    <DllImport("Shell32.dll", EntryPoint:="SHCreateItemFromParsingName")>
    <PreserveSig> Private Shared Function SHCreateItemFromParsingName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
                                                                      <[In]> pbc As IntPtr,
                                                                      <[In], MarshalAs(UnmanagedType.LPStruct)> riid As Guid,
                                                                      <Out> ByRef pUnk As IntPtr) As Integer
    End Function
#End Region

#Region "Enum"
    Public Enum FILEOPENDIALOGOPTIONS As Integer
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

    Public Enum SIGDN As Integer
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

    Private Enum FDAP As Integer
        FDAP_BOTTOM = 0
        FDAP_TOP = 1
    End Enum

    Private Enum FDE_SHAREVIOLATION_RESPONSE As Integer
        FDESVR_DEFAULT = &H0
        FDESVR_ACCEPT = &H1
        FDESVR_REFUSE = &H2
    End Enum
    Private Enum FDE_OVERWRITE_RESPONSE As Integer
        FDEOR_DEFAULT = &H0
        FDEOR_ACCEPT = &H1
        FDEOR_REFUSE = &H2
    End Enum
#End Region

#Region "Const"
    Private Const S_OK As Integer = 0

    Private Const IID_IShellItem As String = "43826d1e-e718-42ee-bc55-a1e261c37bfe"
    Private Const IID_IFileDialog2 As String = "61744fc7-85b5-4791-a9b0-272276309b13"
    Private Const IID_IFileDialogEvents As String = "973510db-7d7f-452b-8975-74a85828d354"

    Private Const CLSID_FileOpenDialog As String = "dc1c5a9c-e88a-4dde-a5a1-60f82a20aef7"
#End Region

#Region "Variable"
    Private m_DisposedValue As Boolean
    Private m_Cookie As Integer
    Private m_FileOpenDialog As Object
#End Region

#Region "Structure"
    Public Structure COMDLG_FILTERSPEC
        <MarshalAs(UnmanagedType.LPWStr)> Dim pszName As String
        <MarshalAs(UnmanagedType.LPWStr)> Dim pszSpec As String
        Sub New(Name As String, Spec As String)
            pszName = Name
            pszSpec = Spec
        End Sub
    End Structure
#End Region

#Region "Public Functions"
    Public Function Show() As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).Show(Form.ActiveForm.Handle) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetFileTypes(FilterSpec As COMDLG_FILTERSPEC()) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetFileTypes(FilterSpec.Length, FilterSpec) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetFileTypeIndex(FileTypeIndex As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetFileTypeIndex(Math.Abs(FileTypeIndex) + 1) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetFileTypeIndex() As Integer
        Dim FileTypeIndex As Integer = -1
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).GetFileTypeIndex(FileTypeIndex) = S_OK Then
                FileTypeIndex -= 1
            End If
        End If
        Return FileTypeIndex
    End Function

    Public Function SetOptions(DialogOptions As FILEOPENDIALOGOPTIONS) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetOptions(DialogOptions) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetOptions() As FILEOPENDIALOGOPTIONS
        Dim DialogOptions As New FILEOPENDIALOGOPTIONS
        If m_FileOpenDialog IsNot Nothing Then
            CType(m_FileOpenDialog, IFileDialog2).GetOptions(DialogOptions)
        End If
        Return DialogOptions
    End Function

    Public Function SetDefaultFolder(Optional Folder As String = Nothing) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If Folder Is Nothing Then Folder = Convert.ToChar(0)
            If SHCreateItemFromParsingName(Folder, IntPtr.Zero,
                                           New Guid(IID_IShellItem),
                                           pIShellItem) = S_OK Then
                If CType(m_FileOpenDialog, IFileDialog2).SetFolder(pIShellItem) = S_OK Then
                    bolRet = True
                End If
                Marshal.Release(pIShellItem)
            End If
        End If
        Return bolRet
    End Function

    Public Function SetFolder(Optional Folder As String = Nothing) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If Folder Is Nothing Then Folder = Convert.ToChar(0)
            If SHCreateItemFromParsingName(Folder, IntPtr.Zero,
                                           New Guid(IID_IShellItem),
                                           pIShellItem) = S_OK Then
                If CType(m_FileOpenDialog, IFileDialog2).SetFolder(pIShellItem) = S_OK Then
                    bolRet = True
                End If
                Marshal.Release(pIShellItem)
            End If
        End If
        Return bolRet
    End Function

    Public Function GetFolder(Optional sign As SIGDN = SIGDN.SIGDN_DESKTOPABSOLUTEPARSING) As String
        Dim strRet As String = String.Empty
        If m_FileOpenDialog IsNot Nothing Then
            Dim psi As IShellItem = Nothing
            If CType(m_FileOpenDialog, IFileDialog2).GetFolder(psi) = S_OK Then
                Dim pszName As IntPtr
                If psi.GetDisplayName(sign, pszName) = S_OK Then
                    strRet = Marshal.PtrToStringUni(pszName)
                    Marshal.FreeCoTaskMem(pszName)
                End If
                Marshal.ReleaseComObject(psi)
            End If
        End If
        Return strRet
    End Function

    '    <PreserveSig> Function GetCurrentSelection(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem) As Integer
    '    <PreserveSig> Function SetFileName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String) As Integer
    '    <PreserveSig> Function GetFileName(<Out, MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String) As Integer
    '    <PreserveSig> Function SetTitle(<[In], MarshalAs(UnmanagedType.LPWStr)> pszTitle As String) As Integer
    '    <PreserveSig> Function SetOkButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer
    '    <PreserveSig> Function SetFileNameLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer

    Public Function GetResult(Optional sign As SIGDN = SIGDN.SIGDN_DESKTOPABSOLUTEPARSING) As String
        Dim strRet As String = String.Empty
        If m_FileOpenDialog IsNot Nothing Then
            Dim ShellItem As IShellItem = Nothing
            If CType(m_FileOpenDialog, IFileDialog2).GetResult(ShellItem) = S_OK Then
                Dim pszName As IntPtr
                If ShellItem.GetDisplayName(sign, pszName) = S_OK Then
                    strRet = Marshal.PtrToStringUni(pszName)
                    Marshal.FreeCoTaskMem(pszName)
                End If
                Marshal.ReleaseComObject(ShellItem)
            End If
        End If
        Return strRet
    End Function


    '    <PreserveSig> Function AddPlace(<[In]> psi As IntPtr, <[In]> fdap As FDAP) As Integer
    '    <PreserveSig> Function SetDefaultExtension(<[In], MarshalAs(UnmanagedType.LPWStr)> pszDefaultExtension As String) As Integer
    Public Function Close() As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).Close = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetClientGuid(guid As Guid) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetClientGuid(guid) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function ClearClientData() As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).ClearClientData = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    '<PreserveSig> Function SetCancelButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
    '<PreserveSig> Function SetNavigationRoot(<[In]> psi As IntPtr) As Integer



#End Region


#Region "Interface IShellItem"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IShellItem)>
    Private Interface IShellItem
        <PreserveSig> Function BindToHandler() As Integer
        <PreserveSig> Function GetParent() As Integer
        <PreserveSig> Function GetDisplayName(<[In]> sigdnName As SIGDN,
                                              <Out> ByRef ppszName As IntPtr) As Integer
        <PreserveSig> Function GetAttributes() As Integer
        <PreserveSig> Function Compare() As Integer
    End Interface
#End Region

#Region "Interface IFileDialog2"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IFileDialog2)>
    Private Interface IFileDialog2
        ' ----==== Interface IModalWindow ====----
        <PreserveSig> Function Show(<[In]> hwndOwner As IntPtr) As Integer

        ' ----==== IFileDialog ====----
        <PreserveSig> Function SetFileTypes(<[In]> cFileTypes As Integer,
                                            <[In], MarshalAs(UnmanagedType.LPArray)> rgFilterSpec As COMDLG_FILTERSPEC()) As Integer
        <PreserveSig> Function SetFileTypeIndex(<[In]> iFileType As Integer) As Integer
        <PreserveSig> Function GetFileTypeIndex(<Out> ByRef piFileType As Integer) As Integer
        <PreserveSig> Function Advise(<[In], MarshalAs(UnmanagedType.Interface)> pfde As IFileDialogEvents,
                                      <Out> ByRef pdwCookie As Integer) As Integer
        <PreserveSig> Function Unadvise(<[In]> dwCookie As Integer) As Integer
        <PreserveSig> Function SetOptions(<[In]> fos As FILEOPENDIALOGOPTIONS) As Integer
        <PreserveSig> Function GetOptions(<Out> ByRef pfos As FILEOPENDIALOGOPTIONS) As Integer
        <PreserveSig> Function SetDefaultFolder(<[In]> psi As IntPtr) As Integer
        <PreserveSig> Function SetFolder(<[In]> psi As IntPtr) As Integer
        <PreserveSig> Function GetFolder(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem) As Integer
        <PreserveSig> Function GetCurrentSelection(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem) As Integer
        <PreserveSig> Function SetFileName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String) As Integer
        <PreserveSig> Function GetFileName(<Out, MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String) As Integer
        <PreserveSig> Function SetTitle(<[In], MarshalAs(UnmanagedType.LPWStr)> pszTitle As String) As Integer
        <PreserveSig> Function SetOkButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer
        <PreserveSig> Function SetFileNameLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function GetResult(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem) As Integer
        <PreserveSig> Function AddPlace(<[In]> psi As IntPtr, <[In]> fdap As FDAP) As Integer
        <PreserveSig> Function SetDefaultExtension(<[In], MarshalAs(UnmanagedType.LPWStr)> pszDefaultExtension As String) As Integer
        <PreserveSig> Function Close() As Integer
        <PreserveSig> Function SetClientGuid(<[In], MarshalAs(UnmanagedType.LPStruct)> guid As Guid) As Integer
        <PreserveSig> Function ClearClientData() As Integer

        'Deprecated. SetFilter is no longer available for use as of Windows 7
        <PreserveSig> Function SetFilter() As Integer

        ' ----==== IFileDialog2 ====----
        <PreserveSig> Function SetCancelButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function SetNavigationRoot(<[In]> psi As IntPtr) As Integer
    End Interface
#End Region

#Region "Interface IFileDialogEvents"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IFileDialogEvents)>
    Private Interface IFileDialogEvents
        <PreserveSig> Function OnFileOk(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer
        <PreserveSig> Function OnFolderChanging(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                                <[In], MarshalAs(UnmanagedType.Interface)> psiFolder As Object) As Integer
        <PreserveSig> Function OnFolderChange(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer
        <PreserveSig> Function OnSelectionChange(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer
        <PreserveSig> Function OnShareViolation(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                                <[In], MarshalAs(UnmanagedType.Interface)> psi As Object,
                                                <Out> ByRef pResponse As Integer) As Integer
        <PreserveSig> Function OnTypeChange(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer
        <PreserveSig> Function OnOverwrite(<[In], MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                           <[In], MarshalAs(UnmanagedType.Interface)> psi As Object,
                                           <Out> ByRef pResponse As Integer) As Integer
    End Interface
#End Region


#Region "Implements IDisposable"
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not m_DisposedValue Then
            If disposing Then
                ' TODO: Verwalteten Zustand (verwaltete Objekte) bereinigen
            End If

            If m_FileOpenDialog IsNot Nothing Then

                If m_Cookie <> 0 Then
                    If CType(m_FileOpenDialog, IFileDialog2).Unadvise(m_Cookie) = S_OK Then
                        m_Cookie = 0
                    End If
                End If

                If Marshal.ReleaseComObject(m_FileOpenDialog) = 0 Then
                    m_FileOpenDialog = Nothing
                End If
            End If

            m_DisposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "Implements IFileDialogEvents"
    Public Function OnFileOk(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer Implements IFileDialogEvents.OnFileOk
        Debug.Print("OnFileOk")
        Return S_OK
    End Function
    Public Function OnFolderChanging(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                     <[In]> <MarshalAs(UnmanagedType.Interface)> psiFolder As Object) As Integer Implements IFileDialogEvents.OnFolderChanging
        Debug.Print("OnFolderChanging")
        Return S_OK
    End Function
    Public Function OnFolderChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer Implements IFileDialogEvents.OnFolderChange
        Debug.Print("OnFolderChange")
        Return S_OK
    End Function
    Public Function OnSelectionChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer Implements IFileDialogEvents.OnSelectionChange
        Debug.Print("OnSelectionChange")
        Return S_OK
    End Function
    Public Function OnShareViolation(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                     <[In]> <MarshalAs(UnmanagedType.Interface)> psi As Object,
                                     <Out> ByRef pResponse As Integer) As Integer Implements IFileDialogEvents.OnShareViolation
        Debug.Print("OnShareViolation")
        Return S_OK
    End Function
    Public Function OnTypeChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object) As Integer Implements IFileDialogEvents.OnTypeChange
        Debug.Print("OnTypeChange")
        Return S_OK
    End Function
    Public Function OnOverwrite(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As Object,
                                <[In]> <MarshalAs(UnmanagedType.Interface)> psi As Object,
                                <Out> ByRef pResponse As Integer) As Integer Implements IFileDialogEvents.OnOverwrite
        Debug.Print("OnOverwrite")
        Return S_OK
    End Function
#End Region

End Class
