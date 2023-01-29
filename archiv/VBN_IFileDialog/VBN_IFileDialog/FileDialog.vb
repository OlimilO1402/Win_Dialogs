Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Public Class FileDialog
    Implements IDisposable
    Implements IFileDialogEvents
    Implements IFileDialogControlEvents

#Region "API"
    <DllImport("Shell32.dll", EntryPoint:="SHCreateItemFromParsingName")>
    <PreserveSig> Private Shared Function SHCreateItemFromParsingName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
                                                                      <[In]> pbc As IntPtr,
                                                                      <[In], MarshalAs(UnmanagedType.LPStruct)> riid As Guid,
                                                                      <Out> ByRef pUnk As IntPtr) As Integer
    End Function
#End Region

#Region "Subs"
    ' Holt das FensterHandle vom Dialog (über die Events vom IFileDialogEvents) einmalig.
    ' Nicht schön, aber reicht. Das FensterHandle kann man dann für ein Subclassing des Dialoges verwenden
    ' um zB. die Controls (API EnumChildWindows) vom Customize anzupassen oder selbst die Controls zu subclassen.
    ' Das Subclassing des Dialoges usw. ist hier aber nicht eingebaut!
    Private Sub GetDialogHwnd(pFileOpenDialog As IntPtr)
        If m_DialogHwnd = IntPtr.Zero Then
            Dim pIOleWindow As IntPtr = IntPtr.Zero
            If Marshal.QueryInterface(pFileOpenDialog, New Guid(IID_IOleWindow), pIOleWindow) = 0 Then
                If pIOleWindow <> IntPtr.Zero Then
                    Dim OleWindow As Object = Marshal.GetObjectForIUnknown(pIOleWindow)
                    If OleWindow IsNot Nothing Then
                        CType(OleWindow, IOleWindow).GetWindow(m_DialogHwnd)
                        Marshal.ReleaseComObject(OleWindow)
                    End If
                    Marshal.Release(pIOleWindow)
                End If
            End If
        End If
    End Sub
#End Region

#Region "Enum"
    Public Enum DialogType As Integer
        OpenFileDialog = 0
        SaveFileDialog = 1
        PicFolderDialog = 2
    End Enum

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

    Public Enum FDAP As Integer
        FDAP_BOTTOM = 0
        FDAP_TOP = 1
    End Enum

    Public Enum FDE_SHAREVIOLATION_RESPONSE As Integer
        FDESVR_DEFAULT = 0
        FDESVR_ACCEPT = 1
        FDESVR_REFUSE = 2
    End Enum

    Public Enum FDE_OVERWRITE_RESPONSE As Integer
        FDEOR_DEFAULT = 0
        FDEOR_ACCEPT = 1
        FDEOR_REFUSE = 2
    End Enum

    Public Enum CDCONTROLSTATEF As Integer
        CDCS_INACTIVE = 0
        CDCS_ENABLED = 1
        CDCS_VISIBLE = 2
        CDCS_ENABLEDVISIBLE = 3
    End Enum
#End Region

#Region "Class"
    Public Sub New(Optional dialogtype As DialogType = DialogType.OpenFileDialog)

        Select Case dialogtype
            Case DialogType.OpenFileDialog, DialogType.PicFolderDialog
                m_FileOpenDialog = Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid(CLSID_FileOpenDialog)))
            Case DialogType.SaveFileDialog
                m_FileOpenDialog = Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid(CLSID_FileSaveDialog)))
        End Select

        If m_FileOpenDialog IsNot Nothing Then

            If dialogtype = DialogType.PicFolderDialog Then Me.SetOptions(Me.GetOptions Or FILEOPENDIALOGOPTIONS.FOS_PICKFOLDERS)

            CType(m_FileOpenDialog, IFileDialog2).Advise(Me, m_Cookie)

            Dim pIFileDialogCustomize As IntPtr = IntPtr.Zero
            If Marshal.QueryInterface(Marshal.GetIUnknownForObject(m_FileOpenDialog), New Guid(IID_IFileDialogCustomize), pIFileDialogCustomize) = 0 Then
                If pIFileDialogCustomize <> IntPtr.Zero Then
                    m_FileDialogCustomize = Marshal.GetObjectForIUnknown(pIFileDialogCustomize)
                    Marshal.Release(pIFileDialogCustomize)
                End If
            End If

        End If
    End Sub

    Protected Overrides Sub Finalize()
        Me.Dispose()
        MyBase.Finalize()
    End Sub
#End Region

#Region "Const"
    Private Const S_OK As Integer = 0

    Private Const IID_IOleWindow As String = "00000114-0000-0000-c000-000000000046"
    Private Const IID_IShellItem As String = "43826d1e-e718-42ee-bc55-a1e261c37bfe"
    Private Const IID_IFileDialog2 As String = "61744fc7-85b5-4791-a9b0-272276309b13"
    Private Const IID_IFileDialogEvents As String = "973510db-7d7f-452b-8975-74a85828d354"
    Private Const IID_IFileDialogCustomize As String = "e6fdd21a-163f-4975-9c8c-a69f1ba37034"
    Private Const IID_IFileDialogControlEvents As String = "36116642-d713-4b97-9b83-7484a9d00433"

    Private Const CLSID_FileOpenDialog As String = "dc1c5a9c-e88a-4dde-a5a1-60f82a20aef7"
    Private Const CLSID_FileSaveDialog As String = "c0b4e2f3-ba21-4773-8dba-335ec946eb8b"
#End Region

#Region "Events"
    Public Event FileOK()
    Public Event FolderChanging(Folder As String)
    Public Event FolderChange()
    Public Event SelectionChange()
    Public Event ShareViolation(Name As String, Response As FDE_SHAREVIOLATION_RESPONSE)
    Public Event TypeChange()
    Public Event Overwrite(Name As String, Response As FDE_OVERWRITE_RESPONSE)

    Public Event ItemSelected(CtlID As Integer, ItemID As Integer)
    Public Event ButtonClicked(CtlID As Integer)
    Public Event CheckButtonToggled(CtlID As Integer, Checked As Boolean)
    Public Event ControlActivating(CtlID As Integer)
#End Region

#Region "Variable"
    Private m_DialogHwnd As IntPtr
    Private m_DisposedValue As Boolean
    Private m_Cookie As Integer
    Private m_FileOpenDialog As Object
    Private m_FileDialogCustomize As Object
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

#Region "Functions"
    ' ----==== Interface IFileDialog2 ====----
    Public Function Show() As Boolean
        Dim bolRet As Boolean = False
        m_DialogHwnd = IntPtr.Zero
        If m_FileOpenDialog IsNot Nothing Then
            Dim ActiveFormHandle = Form.ActiveForm.Handle
            If CType(m_FileOpenDialog, IFileDialog2).Show(ActiveFormHandle) = S_OK Then
                bolRet = True
            End If
        End If
        m_DialogHwnd = IntPtr.Zero
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

    Public Function SetFileTypeIndex(Optional FileTypeIndex As Integer = 0) As Boolean
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

    Public Function SetDefaultFolder(Optional Folder As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If String.IsNullOrEmpty(Folder) Then Folder = Convert.ToChar(0)
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

    Public Function SetFolder(Optional Folder As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If String.IsNullOrEmpty(Folder) Then Folder = Convert.ToChar(0)
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

    Public Function GetCurrentSelection(Optional sign As SIGDN = SIGDN.SIGDN_DESKTOPABSOLUTEPARSING) As String
        Dim strRet As String = String.Empty
        If m_FileOpenDialog IsNot Nothing Then
            Dim psi As IShellItem = Nothing
            If CType(m_FileOpenDialog, IFileDialog2).GetCurrentSelection(psi) = S_OK Then
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

    Public Function SetFileName(Name As String) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetFileName(Name) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetFileName() As String
        Dim strRet As String = String.Empty
        If m_FileOpenDialog IsNot Nothing Then
            CType(m_FileOpenDialog, IFileDialog2).GetFileName(strRet)
        End If
        Return strRet
    End Function

    Public Function SetTitle(Optional Title As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetTitle(Title) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetOkButtonLabel(Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetOkButtonLabel(Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetFileNameLabel(Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetFileNameLabel(Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

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

    Public Function AddPlace(Place As String, Optional fdap As FDAP = FDAP.FDAP_BOTTOM) As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If String.IsNullOrEmpty(Place) Then Place = Convert.ToChar(0)
            If SHCreateItemFromParsingName(Place, IntPtr.Zero,
                                           New Guid(IID_IShellItem),
                                           pIShellItem) = S_OK Then
                If CType(m_FileOpenDialog, IFileDialog2).AddPlace(pIShellItem, fdap) = S_OK Then
                    bolRet = True
                End If
                Marshal.Release(pIShellItem)
            End If
        End If
        Return bolRet
    End Function

    Public Function SetDefaultExtension(Optional DefaultExtension As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetDefaultExtension(DefaultExtension) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

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

    'Deprecated. SetFilter is no longer available for use as of Windows 7
    'Public Function SetFilter() As Boolean
    'End Function

    Public Function SetCancelButtonLabel(Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            If CType(m_FileOpenDialog, IFileDialog2).SetCancelButtonLabel(Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetNavigationRoot(Optional Root As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileOpenDialog IsNot Nothing Then
            Dim pIShellItem As IntPtr
            If String.IsNullOrEmpty(Root) Then Root = Convert.ToChar(0)
            If SHCreateItemFromParsingName(Root, IntPtr.Zero,
                                           New Guid(IID_IShellItem),
                                           pIShellItem) = S_OK Then
                If CType(m_FileOpenDialog, IFileDialog2).SetNavigationRoot(pIShellItem) = S_OK Then
                    bolRet = True
                End If
                Marshal.Release(pIShellItem)
            End If
        End If
        Return bolRet
    End Function


    ' ----==== Interface IFileDialogCustomize ====----
    Public Function EnableOpenDropDown(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).EnableOpenDropDown(CtlID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddMenu(CtlID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddMenu(CtlID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddPushButton(CtlID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddPushButton(CtlID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddComboBox(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddComboBox(CtlID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddRadioButtonList(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddRadioButtonList(CtlID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddCheckButton(CtlID As Integer, Optional Label As String = "", Optional Checked As Boolean = False) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddCheckButton(CtlID, Label, Checked) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddEditBox(CtlID As Integer, Optional Text As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddEditBox(CtlID, Text) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddSeparator(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddSeparator(CtlID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddText(CtlID As Integer, Optional Text As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddText(CtlID, Text) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetControlLabel(CtlID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetControlLabel(CtlID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetControlState(CtlID As Integer) As CDCONTROLSTATEF
        Dim eRet As New CDCONTROLSTATEF
        If m_FileDialogCustomize IsNot Nothing Then
            CType(m_FileDialogCustomize, IFileDialogCustomize).GetControlState(CtlID, eRet)
        End If
        Return eRet
    End Function

    Public Function SetControlState(CtlID As Integer, State As CDCONTROLSTATEF) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetControlState(CtlID, State) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetEditBoxText(CtlID As Integer) As String
        Dim strRet As String = String.Empty
        If m_FileDialogCustomize IsNot Nothing Then
            Dim pStr As IntPtr = IntPtr.Zero
            If CType(m_FileDialogCustomize, IFileDialogCustomize).GetEditBoxText(CtlID, pStr) = S_OK Then
                strRet = Marshal.PtrToStringUni(pStr)
                Marshal.FreeCoTaskMem(pStr)
            End If
        End If
        Return strRet
    End Function

    Public Function SetEditBoxText(CtlID As Integer, Optional Text As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetEditBoxText(CtlID, Text) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetCheckButtonState(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            CType(m_FileDialogCustomize, IFileDialogCustomize).GetCheckButtonState(CtlID, bolRet)
        End If
        Return bolRet
    End Function

    Public Function SetCheckButtonState(CtlID As Integer, Optional Checked As Boolean = False) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetCheckButtonState(CtlID, Checked) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function AddControlItem(CtlID As Integer, ItemID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).AddControlItem(CtlID, ItemID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function RemoveControlItem(CtlID As Integer, ItemID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).RemoveControlItem(CtlID, ItemID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    ' Not implemented!
    'Public Function RemoveAllControlItems(CtlID As Integer) As Boolean
    'End Function

    Public Function GetControlItemState(CtlID As Integer, ItemID As Integer) As CDCONTROLSTATEF
        Dim eRet As New CDCONTROLSTATEF
        If m_FileDialogCustomize IsNot Nothing Then
            CType(m_FileDialogCustomize, IFileDialogCustomize).GetControlItemState(CtlID, ItemID, eRet)
        End If
        Return eRet
    End Function

    Public Function SetControlItemState(CtlID As Integer, ItemID As Integer, State As CDCONTROLSTATEF) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetControlItemState(CtlID, ItemID, State) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function GetSelectedControlItem(CtlID As Integer) As Integer
        Dim iRet As Integer
        If m_FileDialogCustomize IsNot Nothing Then
            CType(m_FileDialogCustomize, IFileDialogCustomize).GetSelectedControlItem(CtlID, iRet)
        End If
        Return iRet
    End Function

    Public Function SetSelectedControlItem(CtlID As Integer, ItemID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetSelectedControlItem(CtlID, ItemID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function StartVisualGroup(CtlID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).StartVisualGroup(CtlID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function EndVisualGroup() As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).EndVisualGroup = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function MakeProminent(CtlID As Integer) As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).MakeProminent(CtlID) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function

    Public Function SetControlItemText(CtlID As Integer, ItemID As Integer, Optional Label As String = "") As Boolean
        Dim bolRet As Boolean = False
        If m_FileDialogCustomize IsNot Nothing Then
            If CType(m_FileDialogCustomize, IFileDialogCustomize).SetControlItemText(CtlID, ItemID, Label) = S_OK Then
                bolRet = True
            End If
        End If
        Return bolRet
    End Function
#End Region

#Region "Interface IOleWindow"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IOleWindow)>
    Private Interface IOleWindow

        <PreserveSig> Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer
        <PreserveSig> Function ContextSensitiveHelp(<[In], MarshalAs(UnmanagedType.Bool)> fEnterMode As Boolean) As Integer

    End Interface
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
        <PreserveSig> Function OnFileOk(<[In]> pfd As IntPtr) As Integer
        <PreserveSig> Function OnFolderChanging(<[In]> pfd As IntPtr,
                                                <[In]> psiFolder As IntPtr) As Integer
        <PreserveSig> Function OnFolderChange(<[In]> pfd As IntPtr) As Integer
        <PreserveSig> Function OnSelectionChange(<[In]> pfd As IntPtr) As Integer
        <PreserveSig> Function OnShareViolation(<[In]> pfd As IntPtr,
                                                <[In]> psi As IntPtr,
                                                <Out> ByRef pResponse As FDE_SHAREVIOLATION_RESPONSE) As Integer
        <PreserveSig> Function OnTypeChange(<[In]> pfd As IntPtr) As Integer
        <PreserveSig> Function OnOverwrite(<[In]> pfd As IntPtr,
                                           <[In]> psi As IntPtr,
                                           <Out> ByRef pResponse As FDE_OVERWRITE_RESPONSE) As Integer
    End Interface
#End Region

#Region "Interface IFileDialogCustomize"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IFileDialogCustomize)>
    Private Interface IFileDialogCustomize
        <PreserveSig> Function EnableOpenDropDown(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function AddMenu(<[In]> dwIDCtl As Integer,
                                       <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function AddPushButton(<[In]> dwIDCtl As Integer,
                                             <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function AddComboBox(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function AddRadioButtonList(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function AddCheckButton(<[In]> dwIDCtl As Integer,
                                              <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String,
                                              <[In], MarshalAs(UnmanagedType.Bool)> bChecked As Boolean) As Integer
        <PreserveSig> Function AddEditBox(<[In]> dwIDCtl As Integer,
                                          <[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer
        <PreserveSig> Function AddSeparator(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function AddText(<[In]> dwIDCtl As Integer,
                                       <[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer
        <PreserveSig> Function SetControlLabel(<[In]> dwIDCtl As Integer,
                                               <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function GetControlState(<[In]> dwIDCtl As Integer,
                                               <Out> ByRef pdwState As CDCONTROLSTATEF) As Integer
        <PreserveSig> Function SetControlState(<[In]> dwIDCtl As Integer,
                                               <[In]> dwState As CDCONTROLSTATEF) As Integer
        <PreserveSig> Function GetEditBoxText(<[In]> dwIDCtl As Integer,
                                              <Out> ByRef ppszText As IntPtr) As Integer
        <PreserveSig> Function SetEditBoxText(<[In]> dwIDCtl As Integer,
                                              <[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer
        <PreserveSig> Function GetCheckButtonState(<[In]> dwIDCtl As Integer,
                                                   <Out, MarshalAs(UnmanagedType.Bool)> ByRef pbChecked As Boolean) As Integer
        <PreserveSig> Function SetCheckButtonState(<[In]> dwIDCtl As Integer,
                                                   <[In], MarshalAs(UnmanagedType.Bool)> bChecked As Boolean) As Integer
        <PreserveSig> Function AddControlItem(<[In]> dwIDCtl As Integer,
                                              <[In]> dwIDItem As Integer,
                                              <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function RemoveControlItem(<[In]> dwIDCtl As Integer,
                                                 <[In]> dwIDItem As Integer) As Integer
        ' Not implemented.
        <PreserveSig> Function RemoveAllControlItems(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function GetControlItemState(<[In]> dwIDCtl As Integer,
                                                   <[In]> dwIDItem As Integer,
                                                   <Out> ByRef pdwState As CDCONTROLSTATEF) As Integer
        <PreserveSig> Function SetControlItemState(<[In]> dwIDCtl As Integer,
                                                   <[In]> dwIDItem As Integer,
                                                   <[In]> dwState As CDCONTROLSTATEF) As Integer
        <PreserveSig> Function GetSelectedControlItem(<[In]> dwIDCtl As Integer,
                                                      <Out> ByRef dwIDItem As Integer) As Integer
        <PreserveSig> Function SetSelectedControlItem(<[In]> dwIDCtl As Integer,
                                                      <[In]> dwIDItem As Integer) As Integer
        <PreserveSig> Function StartVisualGroup(<[In]> dwIDCtl As Integer,
                                                <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
        <PreserveSig> Function EndVisualGroup() As Integer
        <PreserveSig> Function MakeProminent(<[In]> dwIDCtl As Integer) As Integer
        <PreserveSig> Function SetControlItemText(<[In]> dwIDCtl As Integer,
                                                  <[In]> dwIDItem As Integer,
                                                  <[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer
    End Interface
#End Region

#Region "Interface IFileDialogControlEvents"
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(IID_IFileDialogControlEvents)>
    Private Interface IFileDialogControlEvents
        <PreserveSig> Function OnItemSelected(<[In]> pfdc As IntPtr,
                                              <[In]> dwIDCtl As Integer,
                                              <[In]> dwIDItem As Integer) As Integer
        <PreserveSig> Function OnButtonClicked(<[In]> pfdc As IntPtr,
                                               <[In]> dwIDCtl As Integer) As Integer

        <PreserveSig> Function OnCheckButtonToggled(<[In]> pfdc As IntPtr,
                                                    <[In]> dwIDCtl As Integer,
                                                    <[In], MarshalAs(UnmanagedType.Bool)> bChecked As Boolean) As Integer
        <PreserveSig> Function OnControlActivating(<[In]> pfdc As IntPtr,
                                                   <[In]> dwIDCtl As Integer) As Integer

    End Interface
#End Region

#Region "Implements IDisposable"
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not m_DisposedValue Then
            If m_FileOpenDialog IsNot Nothing Then

                If m_FileDialogCustomize IsNot Nothing Then
                    If Marshal.ReleaseComObject(m_FileDialogCustomize) = 0 Then
                        m_FileDialogCustomize = Nothing
                    End If
                End If

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
    Private Function OnFileOk(<[In]> pfd As IntPtr) As Integer Implements IFileDialogEvents.OnFileOk
        GetDialogHwnd(pfd)
        RaiseEvent FileOK()
        Return S_OK
    End Function
    Private Function OnFolderChanging(<[In]> pfd As IntPtr,
                                     <[In]> psiFolder As IntPtr) As Integer Implements IFileDialogEvents.OnFolderChanging
        GetDialogHwnd(pfd)
        Dim oIShellItem As Object = Marshal.GetObjectForIUnknown(psiFolder)
        Dim pszFolder As IntPtr
        If CType(oIShellItem, IShellItem).GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, pszFolder) = S_OK Then
            RaiseEvent FolderChanging(Marshal.PtrToStringUni(pszFolder))
            Marshal.FreeCoTaskMem(pszFolder)
        End If
        Marshal.ReleaseComObject(oIShellItem)
        Return S_OK
    End Function
    Private Function OnFolderChange(<[In]> pfd As IntPtr) As Integer Implements IFileDialogEvents.OnFolderChange
        GetDialogHwnd(pfd)
        RaiseEvent FolderChange()
        Return S_OK
    End Function
    Private Function OnSelectionChange(<[In]> pfd As IntPtr) As Integer Implements IFileDialogEvents.OnSelectionChange
        GetDialogHwnd(pfd)
        RaiseEvent SelectionChange()
        Return S_OK
    End Function
    Private Function OnShareViolation(<[In]> pfd As IntPtr,
                                      <[In]> psi As IntPtr,
                                      <Out> ByRef pResponse As FDE_SHAREVIOLATION_RESPONSE) As Integer Implements IFileDialogEvents.OnShareViolation
        GetDialogHwnd(pfd)
        Dim oIShellItem As Object = Marshal.GetObjectForIUnknown(psi)
        Dim pszName As IntPtr
        If CType(oIShellItem, IShellItem).GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, pszName) = S_OK Then
            RaiseEvent ShareViolation(Marshal.PtrToStringUni(pszName), pResponse)
            Marshal.FreeCoTaskMem(pszName)
        End If
        Marshal.ReleaseComObject(oIShellItem)
        Return S_OK
    End Function
    Private Function OnTypeChange(<[In]> pfd As IntPtr) As Integer Implements IFileDialogEvents.OnTypeChange
        GetDialogHwnd(pfd)
        RaiseEvent TypeChange()
        Return S_OK
    End Function
    Private Function OnOverwrite(<[In]> pfd As IntPtr,
                                 <[In]> psi As IntPtr,
                                 <Out> ByRef pResponse As FDE_OVERWRITE_RESPONSE) As Integer Implements IFileDialogEvents.OnOverwrite
        GetDialogHwnd(pfd)
        Dim oIShellItem As Object = Marshal.GetObjectForIUnknown(psi)
        Dim pszName As IntPtr
        If CType(oIShellItem, IShellItem).GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, pszName) = S_OK Then
            RaiseEvent Overwrite(Marshal.PtrToStringUni(pszName), pResponse)
            Marshal.FreeCoTaskMem(pszName)
        End If
        Marshal.ReleaseComObject(oIShellItem)
        Return S_OK
    End Function
#End Region

#Region "Implements IFileDialogControlEvents"
    Private Function OnItemSelected(<[In]> pfdc As IntPtr,
                                    <[In]> dwIDCtl As Integer,
                                    <[In]> dwIDItem As Integer) As Integer Implements IFileDialogControlEvents.OnItemSelected
        RaiseEvent ItemSelected(dwIDCtl, dwIDItem)
        Return S_OK
    End Function

    Private Function OnButtonClicked(<[In]> pfdc As IntPtr,
                                     <[In]> dwIDCtl As Integer) As Integer Implements IFileDialogControlEvents.OnButtonClicked
        RaiseEvent ButtonClicked(dwIDCtl)
        Return S_OK
    End Function

    Private Function OnCheckButtonToggled(<[In]> pfdc As IntPtr,
                                          <[In]> dwIDCtl As Integer,
                                          <[In]> <MarshalAs(UnmanagedType.Bool)> bChecked As Boolean) As Integer Implements IFileDialogControlEvents.OnCheckButtonToggled
        RaiseEvent CheckButtonToggled(dwIDCtl, bChecked)
        Return S_OK
    End Function

    Private Function OnControlActivating(<[In]> pfdc As IntPtr,
                                         <[In]> dwIDCtl As Integer) As Integer Implements IFileDialogControlEvents.OnControlActivating
        RaiseEvent ControlActivating(dwIDCtl)
        Return S_OK
    End Function
#End Region

End Class
