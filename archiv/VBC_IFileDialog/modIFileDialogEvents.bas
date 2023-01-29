Attribute VB_Name = "modIFileDialogEvents"
Option Explicit

' ----==== Const ====----
Private Const S_OK As Long = &H0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const IID_IFileDialogEvents As String = "{973510DB-7D7F-452B-8975-74A85828D354}"
Private Const IID_IFileDialogControlEvents As String = "{36116642-D713-4b97-9B83-7484A9D00433}"

' ----==== Types ====----
Private Type tIFileDialogEvents
    pVTable As Long
End Type

Private Type tIFileDialogControlEvents
    pVTable As Long
End Type

Private Type tIFileDialogEventsVTable
    VTable(0 To 9) As Long
End Type

Private Type tIFileDialogControlEventsVTable
    VTable(0 To 6) As Long
End Type

' ----==== Kernel32 API-Deklarationen ====----
Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal length As Long)

' ----==== Variablen ====----
Private m_IPtr As Long
Private m_RefCount As Long
Private m_pInterface As Long
Private m_Interface As Object
Private m_IFileDialogEvents As tIFileDialogEvents
Private m_IFileDialogEventsVTable As tIFileDialogEventsVTable
Private m_IFileDialogControlEvents As tIFileDialogControlEvents
Private m_IFileDialogControlEventsVTable As tIFileDialogControlEventsVTable

Public Function IFileDialogEvents(ByVal cIFileDialogEvents As Long) As Long

    ' Pointer auf die Klasse, an die die Events weitergeleitet werden
    m_pInterface = cIFileDialogEvents

    ' ----==== Interface IFileDialogEvents  ====----
    With m_IFileDialogEventsVTable
        
        ' VTable aufbauen
        .VTable(0) = ProcPtr(AddressOf QueryInterface)
        .VTable(1) = ProcPtr(AddressOf AddRef)
        .VTable(2) = ProcPtr(AddressOf Release)
        .VTable(3) = ProcPtr(AddressOf OnFileOk)
        .VTable(4) = ProcPtr(AddressOf OnFolderChanging)
        .VTable(5) = ProcPtr(AddressOf OnFolderChange)
        .VTable(6) = ProcPtr(AddressOf OnSelectionChange)
        .VTable(7) = ProcPtr(AddressOf OnShareViolation)
        .VTable(8) = ProcPtr(AddressOf OnTypeChange)
        .VTable(9) = ProcPtr(AddressOf OnOverwrite)

    End With

    With m_IFileDialogEvents

        ' Pointer auf die VTable
        .pVTable = VarPtr(m_IFileDialogEventsVTable)

    End With

    ' ----==== Interface IFileDialogControlEvents  ====----
    With m_IFileDialogControlEventsVTable

        ' VTable aufbauen
        .VTable(0) = ProcPtr(AddressOf QueryInterface)
        .VTable(1) = ProcPtr(AddressOf AddRef)
        .VTable(2) = ProcPtr(AddressOf Release)
        .VTable(3) = ProcPtr(AddressOf OnItemSelected)
        .VTable(4) = ProcPtr(AddressOf OnButtonClicked)
        .VTable(5) = ProcPtr(AddressOf OnCheckButtonToggled)
        .VTable(6) = ProcPtr(AddressOf OnControlActivating)

    End With

    With m_IFileDialogControlEvents

        ' Pointer auf die VTable
        .pVTable = VarPtr(m_IFileDialogControlEventsVTable)

    End With

    ' Pointer auf das Interface
    m_IPtr = VarPtr(m_IFileDialogEvents)

    IFileDialogEvents = m_IPtr

End Function

Private Function ProcPtr(ByVal ptr As Long) As Long

    ProcPtr = ptr

End Function

' ----==== Interface IUnknown Func ====----
Private Function QueryInterface(ByVal this As Long, ByRef riid As GUID, ByRef _
    pvObj As Long) As Long

    Dim lRet As Long
    
    Select Case UCase$(Guid2String(riid))

    ' wenn nach dem Interface IFileDialogEvents gefragt wird
    Case UCase$(IID_IFileDialogEvents)

        ' dann muss ein AddRef aufgerufen werden.
        Call AddRef(this)
        
        ' Pointer auf das Interface
        pvObj = VarPtr(m_IFileDialogEvents)
        
        ' OK zurück geben, Interface ist vorhanden
        lRet = S_OK

    ' wenn nach dem Interface IFileDialogControlEvents gefragt wird
    ' IFileDialogEvents.QueryInterface(IID_IFileDialogControlEvents)
    Case UCase$(IID_IFileDialogControlEvents)
        
        ' dann muss ein AddRef aufgerufen werden.
        Call AddRef(this)
        
        ' Pointer auf das Interface
        pvObj = VarPtr(m_IFileDialogControlEvents)
        
        ' OK zurück geben, Interface ist vorhanden
        lRet = S_OK
    
    Case Else
        
        ' alles andere ignorieren
        
        pvObj = 0&
        
        lRet = E_NOINTERFACE
    
    End Select

    QueryInterface = lRet

End Function

Private Function AddRef(ByVal this As Long) As Long

    m_RefCount = m_RefCount + 1

    AddRef = m_RefCount

End Function

Private Function Release(ByVal this As Long) As Long
    
    m_RefCount = m_RefCount - 1

    If m_RefCount = 0 Then
    
        ' alles Aufräumen und Freigeben
        m_IFileDialogControlEvents.pVTable = 0&
        Erase m_IFileDialogControlEventsVTable.VTable

        m_IFileDialogEvents.pVTable = 0&
        Erase m_IFileDialogEventsVTable.VTable
        
        m_IPtr = 0&
        
        m_pInterface = 0&
    
    End If
    
    Release = m_RefCount

End Function

' ----==== Interface IFileDialogEvents Func ====----
Private Function OnFileOk(ByVal this As Long, ByVal pfd As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnFileOk(pfd)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnFolderChanging(ByVal this As Long, ByVal pfd As Long, ByVal _
    psiFolder As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnFolderChanging(pfd, psiFolder)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnFolderChange(ByVal this As Long, ByVal pfd As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnFolderChange(pfd)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnSelectionChange(ByVal this As Long, ByVal pfd As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnSelectionChange(pfd)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnShareViolation(ByVal this As Long, ByVal pfd As Long, ByVal psi _
    As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnShareViolation(pfd, psi)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnTypeChange(ByVal this As Long, ByVal pfd As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnTypeChange(pfd)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnOverwrite(ByVal this As Long, ByVal pfd As Long, ByVal psi As _
    Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnOverwrite(pfd, psi)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

' ----==== Interface IFileDialogControlEvents Func ====----
Private Function OnItemSelected(ByVal this As Long, ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal dwIDItem As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnItemSelected(pfdc, dwIDCtl, dwIDItem)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnButtonClicked(ByVal this As Long, ByVal pfdc As Long, ByVal dwIDCtl As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnButtonClicked(pfdc, dwIDCtl)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnCheckButtonToggled(ByVal this As Long, ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal bChecked As Boolean) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnCheckButtonToggled(pfdc, dwIDCtl, bChecked)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If
    
End Function

Private Function OnControlActivating(ByVal this As Long, ByVal pfdc As Long, ByVal dwIDCtl As Long) As Long

    If m_pInterface <> 0& Then

        ' Objekt vom Pointer
        Call CopyMemory(m_Interface, m_pInterface, 4)
            
        ' Funktion in der Klasse aufrufen
        Call m_Interface.OnControlActivating(pfdc, dwIDCtl)

        ' Objekt löschen
        Call CopyMemory(m_Interface, 0&, 4)

    End If

End Function

