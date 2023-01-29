VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIEnumShellItems"
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
Private Const CLSID_EnumShellItems As String = ""
Private Const IID_IEnumShellItems As String = "{70629033-e363-4a28-a567-0db78006e6d7}"
    
Private Enum E_vtbOffsets
    '/*** IUnknown methods ***/
    vtb_QueryInterface
    vtb_AddRef
    vtb_Release
    '/*** IEnumShellItems methods ***/
    vtb_Fetch 'Next
    vtb_Skip
    vtb_Reset
    vtb_Clone
End Enum

Private m_initialized As Boolean
Private m_Interface As Interface_Data

' ----==== Interface IEnumShellItems Func ====----
Public Function Fetch(ByVal celt As Long, ByRef rgelt() As Long, ByRef _
    pceltFetched As Long) As Boolean

    Dim lItem As Long
    Dim lCount As Long
    Dim bRet As Boolean
    Dim prgelt As Long

    bRet = False

    If celt > 0 Then

        lCount = celt - 1
        ReDim rgelt(lCount)
        prgelt = VarPtr(rgelt(0))

        If x_OleInvoke(vtb_Fetch, celt, prgelt, pceltFetched) Then

            bRet = True

        End If

    End If

    Fetch = bRet

End Function

Public Function Skip(ByVal celt As Long) As Boolean

    Dim bRet As Boolean

    If x_OleInvoke(vtb_Skip, celt) Then

        bRet = True

    End If

    Skip = bRet

End Function

Public Function Reset() As Boolean

    Dim bRet As Boolean

    If x_OleInvoke(vtb_Reset) Then

        bRet = True

    End If

    Reset = bRet

End Function

Public Function Clone() As Long

    Dim lRet As Long
    Dim penum As Long
    Dim ppenum As Long

    ppenum = VarPtr(penum)

    If x_OleInvoke(vtb_Clone, ppenum) Then

        lRet = penum

    End If

    Clone = lRet

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
            ".clsIEnumShellItems", .etx
            
        .ecd = 0
        .etx = ""
        
    End With
    
End Sub

