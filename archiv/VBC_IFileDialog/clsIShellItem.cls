VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIShellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2017

' Komplett

Option Explicit

Private Const ERR_BASE As Long = 40670
Private Const CLSID_ShellItem As String = "{9ac9fbe1-e0a2-4ad6-b4ee-e212013ea917}"
Private Const IID_IShellItem As String = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"

Private Enum E_vtbOffsets
    '/*** IUnknown methods ***/
    vtb_QueryInterface
    vtb_AddRef
    vtb_Release
    '/*** IShellItem methods ***/
    vtb_BindToHandler
    vtb_GetParent
    vtb_GetDisplayName
    vtb_GetAttributes
    vtb_Compare
End Enum

Private m_initialized As Boolean
Private m_Interface As Interface_Data

' ----==== Interface IShellItem Func ====----
Public Function BindToHandler(ByVal pbc As Long, ByVal rbhid As Long, ByVal riid _
    As Long) As Long

    Dim lRet As Long
    Dim ppvOut As Long
    Dim pppvOut As Long

    pppvOut = VarPtr(ppvOut)

    If x_OleInvoke(vtb_BindToHandler, pbc, rbhid, riid, pppvOut) Then

        lRet = ppvOut

    End If

    BindToHandler = lRet

End Function

Public Function GetParent() As Long

    Dim lRet As Long
    Dim ppsi As Long
    Dim pppsi As Long

    pppsi = VarPtr(ppsi)

    If x_OleInvoke(vtb_GetParent, pppsi) Then

        lRet = ppsi

    End If

    GetParent = lRet

End Function

Public Function GetDisplayName(ByVal sigdnName As SIGDN) As String

    Dim sRet As String
    Dim pszName As Long
    Dim ppszName As Long

    ppszName = VarPtr(pszName)

    If x_OleInvoke(vtb_GetDisplayName, sigdnName, ppszName) Then

        sRet = GetStringFromPointer(pszName)

    End If

    GetDisplayName = sRet

End Function

Public Function GetAttributes(ByVal sfgaoMask As Long) As Long

    Dim lRet As Long
    Dim psfgaoAttribs As Long
    Dim ppsfgaoAttribs As Long

    ppsfgaoAttribs = VarPtr(psfgaoAttribs)

    If x_OleInvoke(vtb_GetAttributes, sfgaoMask, ppsfgaoAttribs) Then

        lRet = psfgaoAttribs

    End If

    GetAttributes = lRet

End Function

Public Function Compare(ByVal psi As Long, ByVal hint As SICHINTF) As Long
    
    Dim lRet As Long
    Dim piOrder As Long
    Dim ppiOrder As Long
    
    ppiOrder = VarPtr(piOrder)
    
    If x_OleInvoke(vtb_Compare, psi, hint, ppiOrder) Then

        lRet = piOrder

    End If

    Compare = lRet
    
End Function

' ----==== Interface Helper Func ====----
Public Sub Init(ByVal pInterface As Long)

    m_Interface.owner = ObjPtr(Me)
    m_Interface.RaiseErrors = False
    m_Interface.ifc = pInterface
    m_initialized = True

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

        If .ecd Then Err.Raise ERR_BASE + .ecd, App.EXEName & ".clsIShellItem", .etx

        .ecd = 0
        .etx = ""

    End With

End Sub

