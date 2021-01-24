VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenFolderDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Many thanks to Frank Schüler for the great efforts he makes every day for ActiveVB

Private Const S_OK As Long = &H0&
Private Const FOS_PICKFOLDERS = &H20
Private Const SIGDN_FILESYSPATH = &H80058000
Private Const CC_STDCALL As Long = &H4&
Private Const IID_Release As Long = &H8&
Private Const CLSCTX_INPROC As Long = &H1&
Private Const CLSID_FileOpenDialog As String = "{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}"
Private Const IID_IFileDialog As String = "{42F85136-DB7E-439C-85F1-E4075D135FC8}"
Private Const IID_IShellItem As String = "{43826D1E-E718-42EE-BC55-A1E261C37BFE}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IModalWindow
    Show_ = 3
    
    ' IFileDialog
    SetOptions_ = 9
    GetOptions_ = 10
    SetFolder_ = 12
    SetTitle_ = 17
    GetResult_ = 20

    'IShellItem
    GetDisplayName_ = 5

End Enum

' ----==== Type ====----
Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7) As Byte
End Type

' ----==== Kernel32.dll Deklarationen ====----
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal ByteLen As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

' ----==== Shell32.dll Deklarationen ====----
Private Declare Function SHCreateItemFromParsingName Lib "Shell32" (ByVal pszPath As Long, ByVal pbc As Long, ByVal riid As Long, ByRef ppv As Long) As Long

' ----==== Ole32.dll Deklarationen ====----
Private Declare Function CLSIDFromString Lib "Ole32" (ByVal pString As Long, ByRef pCLSID As GUID) As Long
Private Declare Function CoCreateInstance Lib "Ole32" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Long) As Long
Private Declare Sub CoTaskMemFree Lib "Ole32" (ByVal hMem As Long)

' ----==== Oleaut32.dll Deklarationen ====----
Private Declare Function DispCallFunc Lib "Oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, ByVal prgvt As Any, ByVal prgpvarg As Any, ByRef pvargResult As Variant) As Long

' ----==== Variablen ====----
Private m_pIPickFolderDialog As Long

Private Sub Class_Initialize()
    Dim lngOptions As Long
    If CoCreateInstance(VarPtr(Str2Guid(CLSID_FileOpenDialog)), 0&, CLSCTX_INPROC, VarPtr(Str2Guid(IID_IFileDialog)), m_pIPickFolderDialog) = S_OK Then
        lngOptions = Me.Options
        lngOptions = lngOptions Or FOS_PICKFOLDERS
        Me.Options = lngOptions
    End If
End Sub

Private Sub Class_Terminate()
    Release m_pIPickFolderDialog
End Sub

Private Sub Release(ByRef pInterface As Long)
    Dim vRet As Variant
    If pInterface = 0& Then Exit Sub
    If DispCallFunc(pInterface, IID_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, vRet) = S_OK Then
        pInterface = 0&
    End If
End Sub

Public Property Let Title(ByVal sTitle As String)
    If Len(sTitle) = 0 Then Exit Property
    If Invoke(m_pIPickFolderDialog, SetTitle_, StrPtr(sTitle)) = S_OK Then
        'bolRet = True
    End If
End Property

Public Function ShowDialog(Optional ByVal hwndOwner As Long) As VbMsgBoxResult
    If Invoke(m_pIPickFolderDialog, Show_, hwndOwner) = S_OK Then
        'bolRet = True
        ShowDialog = vbOK
    Else
        ShowDialog = vbCancel
    End If
End Function

Public Property Let Folder(ByVal sStartFolder As String)
    'Create an IShellItem-object and destroy it afterwards
    Dim pIShellItem As Long
    If SHCreateItemFromParsingName(StrPtr(sStartFolder), 0&, VarPtr(Str2Guid(IID_IShellItem)), pIShellItem) = S_OK Then
        If Invoke(m_pIPickFolderDialog, SetFolder_, pIShellItem) = S_OK Then
            'bolRet = True
        End If
        Call Release(pIShellItem)
    End If
End Property

Public Property Get Folder() As String
    Dim strRet As String
    Dim pszName As Long
    Dim pIShellItem As Long
    If Invoke(m_pIPickFolderDialog, GetResult_, VarPtr(pIShellItem)) = S_OK Then
        If Invoke(pIShellItem, GetDisplayName_, SIGDN_FILESYSPATH, VarPtr(pszName)) = S_OK Then
            strRet = Ptr2String(pszName)
        End If
        Call Release(pIShellItem)
    End If
    Folder = strRet
End Property

Public Property Get Options() As Long
    If Invoke(m_pIPickFolderDialog, vtb_Interfaces.GetOptions_, VarPtr(Options)) = S_OK Then
        'bolRet = True
    End If
End Property

Public Property Let Options(ByVal Value As Long) 'As Boolean
    If Invoke(m_pIPickFolderDialog, vtb_Interfaces.SetOptions_, Value) = S_OK Then
        'bolRet = True
    End If
End Property

Private Function Invoke(ByVal pInterface As Long, ByVal vtb As vtb_Interfaces, ParamArray var()) As Variant
    If pInterface = 0& Then Exit Function
    Invoke = OleInvoke(pInterface, vtb, var)
    'End If
End Function

Private Function OleInvoke(ByVal pInterface As Long, ByVal lCmd As Long, ParamArray aParam()) As Variant
    If pInterface = 0& Then Exit Function
    Dim olePtr(10) As Long
    Dim oleTyp(10) As Integer
    If UBound(aParam) >= 0 Then
        Dim vParam As Variant:  vParam = aParam
        If IsArray(vParam) Then vParam = vParam(0)
        Dim lCount As Long: lCount = UBound(vParam)
        Dim lItem As Long
        For lItem = 0 To lCount
            oleTyp(lItem) = VarType(vParam(lItem))
            olePtr(lItem) = VarPtr(vParam(lItem))
        Next
    End If
    Dim vRet As Variant
    If DispCallFunc(pInterface, lCmd * 4, CC_STDCALL, vbLong, lItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), vRet) <> S_OK Then
        Debug.Print "Error calling the Interface-function: " & Vtb2FunctionName(lCmd)
    End If
    OleInvoke = vRet
End Function

Private Function Str2Guid(ByVal str As String) As GUID
    CLSIDFromString StrPtr(str), Str2Guid
End Function

Private Function Ptr2String(ByVal lpStrPointer As Long) As String
    Dim lLen As Long
    Dim bBuffer() As Byte
    If lpStrPointer = 0 Then Exit Function
    lLen = lstrlenW(lpStrPointer) * 2& - 1&
    If lLen > 0& Then
        ReDim bBuffer(lLen)
        RtlMoveMemory bBuffer(0), ByVal lpStrPointer, lLen
        Call CoTaskMemFree(lpStrPointer)
        Ptr2String = bBuffer
    End If
End Function

Private Function Vtb2FunctionName(f As vtb_Interfaces) As String
    Dim s As String
    Select Case f
    ' IModalWindow
    Case Show_:            s = "Show_"           ' "ShowDialog"
    'IShellItem 'lean
    Case GetDisplayName_:  s = "GetDisplayName_" ' "Property Get DisplayName"
    ' IFileDialog
    Case SetOptions_ = 9:  s = "SetOptions_"     ' "Property Let Options"
    Case GetOptions_ = 10: s = "GetOptions_"     ' "Property Get Options"
    Case SetFolder_ = 12:  s = "SetFolder_"      ' "Property Let Folder"
    Case SetTitle_ = 17:   s = "SetTitle_"       ' "Property Let Title"
    Case GetResult_ = 20:  s = "GetResult_"      ' "Property Get Folder"
    End Select
    Vtb2FunctionName = s
End Function
