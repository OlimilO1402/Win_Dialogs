VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIFileDialogCustomize"
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
Private Const CLSID_FileDialogCustomize As String = ""
Private Const IID_IFileDialogCustomize As String = "{e6fdd21a-163f-4975-9c8c-a69f1ba37034}"

Private Enum E_vtbOffsets
    ' /*** IUnknown methods ***/
    vtb_QueryInterface
    vtb_AddRef
    vtb_Release
    ' /*** IFileDialogCustomize ***/
    vtb_EnableOpenDropDown
    vtb_AddMenu
    vtb_AddPushButton
    vtb_AddComboBox
    vtb_AddRadioButtonList
    vtb_AddCheckButton
    vtb_AddEditBox
    vtb_AddSeparator
    vtb_AddText
    vtb_SetControlLabel
    vtb_GetControlState
    vtb_SetControlState
    vtb_GetEditBoxText
    vtb_SetEditBoxText
    vtb_GetCheckButtonState
    vtb_SetCheckButtonState
    vtb_AddControlItem
    vtb_RemoveControlItem
    vtb_RemoveAllControlItems
    vtb_GetControlItemState
    vtb_SetControlItemState
    vtb_GetSelectedControlItem
    vtb_SetSelectedControlItem
    vtb_StartVisualGroup
    vtb_EndVisualGroup
    vtb_MakeProminent
    vtb_SetControlItemText
End Enum

Private m_initialized As Boolean
Private m_Interface As Interface_Data

' ----==== Interface IFileDialogCustomize Func ====----
Public Function EnableOpenDropDown(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_EnableOpenDropDown, dwIDCtl) Then

        bRet = True

    End If

    EnableOpenDropDown = bRet

End Function

Public Function AddMenu(ByVal dwIDCtl As Long, ByVal szLabel As String) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_AddMenu, dwIDCtl, pszLabel) Then

        bRet = True

    End If

    AddMenu = bRet

End Function

Public Function AddPushButton(ByVal dwIDCtl As Long, ByVal szLabel As String) As _
    Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_AddPushButton, dwIDCtl, pszLabel) Then

        bRet = True

    End If

    AddPushButton = bRet

End Function

Public Function AddComboBox(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_AddComboBox, dwIDCtl) Then

        bRet = True

    End If

    AddComboBox = bRet

End Function

Public Function AddRadioButtonList(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_AddRadioButtonList, dwIDCtl) Then

        bRet = True

    End If

    AddRadioButtonList = bRet

End Function

Public Function AddCheckButton(ByVal dwIDCtl As Long, ByVal szLabel As String, _
    ByVal bChecked As Boolean) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_AddCheckButton, dwIDCtl, pszLabel, bChecked) Then

        bRet = True

    End If

    AddCheckButton = bRet

End Function

Public Function AddEditBox(ByVal dwIDCtl As Long, ByVal szText As String) As Boolean

    Dim bRet As Boolean
    Dim pszText As Long

    bRet = False

    pszText = StrPtr(szText)

    If x_OleInvoke(vtb_AddEditBox, dwIDCtl, pszText) Then

        bRet = True

    End If

    AddEditBox = bRet

End Function

Public Function AddSeparator(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_AddSeparator, dwIDCtl) Then

        bRet = True

    End If

    AddSeparator = bRet

End Function

Public Function AddText(ByVal dwIDCtl As Long, ByVal szText As String) As Boolean

    Dim bRet As Boolean
    Dim pszText As Long

    bRet = False

    pszText = StrPtr(szText)

    If x_OleInvoke(vtb_AddText, dwIDCtl, pszText) Then

        bRet = True

    End If

    AddText = bRet

End Function

Public Function SetControlLabel(ByVal dwIDCtl As Long, ByVal szLabel As String) As _
    Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_SetControlLabel, dwIDCtl, pszLabel) Then

        bRet = True

    End If

    SetControlLabel = bRet

End Function

Public Function GetControlState(ByVal dwIDCtl As Long) As CDCONTROLSTATEF

    Dim lRet As Long
    Dim dwState As Long
    Dim pdwState As Long

    pdwState = VarPtr(dwState)

    If x_OleInvoke(vtb_GetControlState, dwIDCtl, pdwState) Then

        lRet = dwState

    End If

    GetControlState = lRet

End Function

Public Function SetControlState(ByVal dwIDCtl As Long, ByVal dwState As _
    CDCONTROLSTATEF) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetControlState, dwIDCtl, dwState) Then

        bRet = True

    End If

    SetControlState = bRet

End Function

Public Function GetEditBoxText(ByVal dwIDCtl As Long) As String

    Dim sRet As String
    Dim pszText As Long
    Dim ppszText As Long

    ppszText = StrPtr(pszText)

    If x_OleInvoke(vtb_GetEditBoxText, dwIDCtl, ppszText) Then

        sRet = GetStringFromPointer(pszText)

    End If

    GetEditBoxText = sRet

End Function

Public Function SetEditBoxText(ByVal dwIDCtl As Long, ByVal szText As String) As _
    Boolean

    Dim bRet As Boolean
    Dim pszText As Long

    bRet = False

    pszText = StrPtr(szText)

    If x_OleInvoke(vtb_SetEditBoxText, dwIDCtl, pszText) Then

        bRet = True

    End If

    SetEditBoxText = bRet

End Function

Public Function GetCheckButtonState(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean
    Dim bChecked As Boolean
    Dim pbChecked As Long

    pbChecked = VarPtr(bChecked)

    If x_OleInvoke(vtb_GetCheckButtonState, dwIDCtl, pbChecked) Then

        bRet = bChecked

    End If

    GetCheckButtonState = bRet

End Function

Public Function SetCheckButtonState(ByVal dwIDCtl As Long, ByVal bChecked As _
    Boolean) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetCheckButtonState, dwIDCtl, bChecked) Then

        bRet = True

    End If

    SetCheckButtonState = bRet

End Function

Public Function AddControlItem(ByVal dwIDCtl As Long, ByVal dwIDItem As Long, _
    ByVal szLabel As String) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_AddControlItem, dwIDCtl, dwIDItem, pszLabel) Then

        bRet = True

    End If

    AddControlItem = bRet

End Function

Public Function RemoveControlItem(ByVal dwIDCtl As Long, ByVal dwIDItem As Long) _
    As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_RemoveControlItem, dwIDCtl, dwIDItem) Then

        bRet = True

    End If

    RemoveControlItem = bRet

End Function

Public Function RemoveAllControlItems(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_RemoveAllControlItems, dwIDCtl) Then

        bRet = True

    End If

    RemoveAllControlItems = bRet

End Function

Public Function GetControlItemState(ByVal dwIDCtl As Long, ByVal dwIDItem As Long) _
    As CDCONTROLSTATEF

    Dim lRet As Long
    Dim dwState As Long
    Dim pdwState As Long

    pdwState = VarPtr(dwState)

    If x_OleInvoke(vtb_GetControlItemState, dwIDCtl, dwIDItem, pdwState) Then

        lRet = dwState

    End If

    GetControlItemState = lRet

End Function

Public Function SetControlItemState(ByVal dwIDCtl As Long, ByVal dwIDItem As Long, _
    ByVal dwState As CDCONTROLSTATEF) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetControlItemState, dwIDCtl, dwIDItem, dwState) Then

        bRet = True

    End If

    SetControlItemState = bRet

End Function

Public Function GetSelectedControlItem(ByVal dwIDCtl As Long) As Long

    Dim lRet As Long
    Dim dwIDItem As Long
    Dim pdwIDItem As Long

    pdwIDItem = VarPtr(dwIDItem)

    If x_OleInvoke(vtb_GetSelectedControlItem, dwIDCtl, pdwIDItem) Then

        lRet = dwIDItem

    End If

    GetSelectedControlItem = lRet

End Function

Public Function SetSelectedControlItem(ByVal dwIDCtl As Long, ByVal dwIDItem As _
    Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_SetSelectedControlItem, dwIDCtl, dwIDItem) Then

        bRet = True

    End If

    SetSelectedControlItem = bRet

End Function

Public Function StartVisualGroup(ByVal dwIDCtl As Long, ByVal szLabel As String) _
    As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_StartVisualGroup, dwIDCtl, pszLabel) Then

        bRet = True

    End If

    StartVisualGroup = bRet

End Function

Public Function EndVisualGroup() As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_EndVisualGroup) Then

        bRet = True

    End If

    EndVisualGroup = bRet

End Function

Public Function MakeProminent(ByVal dwIDCtl As Long) As Boolean

    Dim bRet As Boolean

    bRet = False

    If x_OleInvoke(vtb_MakeProminent, dwIDCtl) Then

        bRet = True

    End If

    MakeProminent = bRet

End Function

Public Function SetControlItemText(ByVal dwIDCtl As Long, ByVal dwIDItem As Long, _
    ByVal szLabel As String) As Boolean

    Dim bRet As Boolean
    Dim pszLabel As Long

    bRet = False

    pszLabel = StrPtr(szLabel)

    If x_OleInvoke(vtb_SetControlItemText, dwIDCtl, dwIDItem, pszLabel) Then

        bRet = True

    End If

    SetControlItemText = bRet

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

        If .ecd Then Err.Raise ERR_BASE + .ecd, App.EXEName & _
            ".clsIFileDialogCustomize", .etx

        .ecd = 0
        .etx = ""

    End With

End Sub

