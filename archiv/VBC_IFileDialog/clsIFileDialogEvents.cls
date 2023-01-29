VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIFileDialogEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ----==== Events IFileDialogEvents ====----
Public Event OnFileOk(ByVal pfd As Long)
Public Event OnFolderChanging(ByVal pfd As Long, ByVal psiFolder As Long)
Public Event OnFolderChange(ByVal pfd As Long)
Public Event OnSelectionChange(ByVal pfd As Long)
Public Event OnShareViolation(ByVal pfd As Long, ByVal psi As Long)
Public Event OnTypeChange(ByVal pfd As Long)
Public Event OnOverwrite(ByVal pfd As Long, ByVal psi As Long)

' ----==== Events IFileDialogControlEvents ====----
Public Event OnItemSelected(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal dwIDItem As Long)
Public Event OnButtonClicked(ByVal pfdc As Long, ByVal dwIDCtl As Long)
Public Event OnCheckButtonToggled(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal bChecked As Boolean)
Public Event OnControlActivating(ByVal pfdc As Long, ByVal dwIDCtl As Long)

' ----==== Variablen ====----
Private pIFileDialogEvents As Long

Private Sub Class_Initialize()

    ' Interface IFileDialogEvents initialisieren
    pIFileDialogEvents = modIFileDialogEvents.IFileDialogEvents(ObjPtr(Me))

End Sub

Private Sub Class_Terminate()
    
    ' ist ein Pointer auf das Interface vorhanden
    If pIFileDialogEvents <> 0 Then
    
        ' ruft Release des Interfaces auf
        'Call CoTaskMemFree(pIFileDialogEvents)
    
    End If
    
End Sub

' ----==== liefert den Pointer auf IFileDialogEvents zurück ====----
Public Function GetIPtr() As Long

    GetIPtr = pIFileDialogEvents

End Function

' ----==== Raise Events IFileDialogEvents ====----
Public Sub OnFileOk(ByVal pfd As Long)

    RaiseEvent OnFileOk(pfd)

End Sub

Public Sub OnFolderChanging(ByVal pfd As Long, ByVal psiFolder As Long)

    RaiseEvent OnFolderChanging(pfd, psiFolder)

End Sub

Public Sub OnFolderChange(ByVal pfd As Long)

    RaiseEvent OnFolderChange(pfd)

End Sub

Public Sub OnSelectionChange(ByVal pfd As Long)

    RaiseEvent OnSelectionChange(pfd)

End Sub

Public Sub OnShareViolation(ByVal pfd As Long, ByVal psi As Long)

    RaiseEvent OnShareViolation(pfd, psi)

End Sub

Public Sub OnTypeChange(ByVal pfd As Long)

    RaiseEvent OnTypeChange(pfd)

End Sub

Public Sub OnOverwrite(ByVal pfd As Long, ByVal psi As Long)

    RaiseEvent OnOverwrite(pfd, psi)

End Sub

' ----==== Raise Events IFileDialogControlEvents ====----
Public Sub OnItemSelected(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal _
    dwIDItem As Long)

    RaiseEvent OnItemSelected(pfdc, dwIDCtl, dwIDItem)

End Sub

Public Sub OnButtonClicked(ByVal pfdc As Long, ByVal dwIDCtl As Long)

    RaiseEvent OnButtonClicked(pfdc, dwIDCtl)

End Sub

Public Sub OnCheckButtonToggled(ByVal pfdc As Long, ByVal dwIDCtl As Long, ByVal _
    bChecked As Boolean)

    RaiseEvent OnCheckButtonToggled(pfdc, dwIDCtl, bChecked)

End Sub

Public Sub OnControlActivating(ByVal pfdc As Long, ByVal dwIDCtl As Long)

    RaiseEvent OnControlActivating(pfdc, dwIDCtl)

End Sub

