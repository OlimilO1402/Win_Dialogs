VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIShellItemArray"
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
Private Const CLSID_ShellItemArray As String = ""
Private Const IID_IShellItemArray As String = "{B63EA76D-1F85-456F-A19C-48159EFA858B}"
    
Private Enum E_vtbOffsets
    '/*** IUnknown methods ***/
    vtb_QueryInterface
    vtb_AddRef
    vtb_Release
    '/*** IShellItemArray methods ***/
    vtb_BindToHandler
    vtb_GetPropertyStore
    vtb_GetPropertyDescriptionList
    vtb_GetAttributes
    vtb_GetCount
    vtb_GetItemAt
    vtb_EnumItems
End Enum

Private m_initialized As Boolean
Private m_Interface As Interface_Data

' ----==== Interface IShellItemArray Func ====----
Public Function BindToHandler(ByVal pbc As Long, ByVal rbhid As Long, ByVal riid _
    As Long) As Long

    Dim pRet As Long
    Dim ppvOut As Long
    Dim pppvOut As Long

    pppvOut = VarPtr(ppvOut)

    If x_OleInvoke(vtb_BindToHandler, pbc, rbhid, riid, pppvOut) Then

        pRet = ppvOut

    End If

    BindToHandler = pRet

End Function

Public Function GetPropertyStore(ByVal flags As GETPROPERTYSTOREFLAGS, ByVal riid As Long) As Long

    Dim pRet As Long
    Dim ppv As Long
    Dim pppv As Long

    pppv = VarPtr(ppv)

    If x_OleInvoke(vtb_GetPropertyStore, flags, riid, pppv) Then

        pRet = ppv

    End If

    GetPropertyStore = pRet

End Function

Friend Function GetPropertyDescriptionList(ByRef keyType As PROPERTYKEY, ByVal riid As Long) As Long

    Dim pRet As Long
    Dim ppv As Long
    Dim pppv As Long
    Dim pkeyType As Long
    
    pppv = VarPtr(ppv)
    pkeyType = VarPtr(keyType)
    
    If x_OleInvoke(vtb_GetPropertyDescriptionList, pkeyType, riid, pppv) Then

        pRet = ppv

    End If

    GetPropertyDescriptionList = pRet

End Function

Public Function GetAttributes(ByVal AttribFlags As SIATTRIBFLAGS, ByVal sfgaoMask As SFGAOF) As SFGAOF

    Dim lRet As Long
    Dim psfgaoAttribs As Long
    Dim ppsfgaoAttribs As Long

    ppsfgaoAttribs = VarPtr(psfgaoAttribs)

    If x_OleInvoke(vtb_GetAttributes, AttribFlags, sfgaoMask, ppsfgaoAttribs) Then

        lRet = psfgaoAttribs

    End If

    GetAttributes = lRet

End Function

Public Function GetCount() As Long

    Dim lRet As Long
    Dim pdwNumItems As Long
    Dim ppdwNumItems As Long

    ppdwNumItems = VarPtr(pdwNumItems)

    If x_OleInvoke(vtb_GetCount, ppdwNumItems) Then

        lRet = pdwNumItems

    End If

    GetCount = lRet

End Function

Public Function GetItemAt(ByVal dwIndex As Long) As Long

    Dim pRet As Long
    Dim ppsi As Long
    Dim pppsi As Long

    pppsi = VarPtr(ppsi)

    If x_OleInvoke(vtb_GetItemAt, dwIndex, pppsi) Then

        pRet = ppsi

    End If

    GetItemAt = pRet

End Function

Public Function EnumItems() As Long

    Dim pRet As Long
    Dim ppenumShellItems As Long
    Dim pppenumShellItems As Long

    pppenumShellItems = VarPtr(ppenumShellItems)

    If x_OleInvoke(vtb_EnumItems, pppenumShellItems) Then

        pRet = ppenumShellItems

    End If

    EnumItems = pRet

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

Private Function x_OleInvoke(vtb As E_vtbOffsets, ParamArray var()) As _
    Boolean
    
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
            ".clsIShellItemArray", .etx
            
        .ecd = 0
        .etx = ""
        
    End With
    
End Sub
