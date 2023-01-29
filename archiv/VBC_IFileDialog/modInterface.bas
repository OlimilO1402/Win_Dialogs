Attribute VB_Name = "modInterface"
' Universal Module für alle Interface-Klassen
' Ursprünglich glaub von Udo Schmidt (ActiveVB)

Option Explicit

' ----==== Const ====----
Private Const S_OK As Long = &H0
Private Const CLSCTX_INPROC As Long = &H1
Private Const CC_STDCALL As Long = &H4
Private Const IID_Release As Long = &H8

' ----==== Interface Error Code ====----
Public Enum Interface_errCodes
    ecd_None                    ' no error
    ecd_InvalidCall             ' invalid function call
    ecd_OleConvert              ' could not convert classid
    ecd_InitInterface           ' could not convert interface id
    ecd_OleInvoke               ' could not invoke interface function
End Enum

' ----==== Holds the Interface Data ====----
Public Type Interface_Data
    ifc As Long                 ' Interface-Pointer
    ecd As Interface_errCodes   ' Fehlercode
    etx As String               ' Fehlertext
    owner As Long               ' Pointer zur Klasse
    RaiseErrors As Boolean      ' Fehler auslösen?
End Type

' ----==== Kernel32 API-Deklarationen ====----
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
                    ByRef hpvDest As Any, _
                    ByRef hpvSource As Any, _
                    ByVal cbCopy As Long)

' ----==== Ole32 API-Deklarationen ====----
Private Declare Function CLSIDFromString Lib "ole32" ( _
                         ByVal lpszProgID As Long, _
                         ByRef pCLSID As Any) As Long

Private Declare Function CoCreateInstance Lib "ole32" ( _
                         ByRef rclsid As Any, _
                         ByVal pUnkOuter As Long, _
                         ByVal dwClsContext As Long, _
                         ByRef riid As Any, _
                         ByRef ppv As Long) As Long

' ----==== OleAut32 API-Deklarationen ====----
Private Declare Sub DispCallFunc Lib "OleAut32" ( _
                    ByVal ppv As Long, _
                    ByVal oVft As Long, _
                    ByVal cc As Long, _
                    ByVal rtTYP As VbVarType, _
                    ByVal paCNT As Long, _
                    ByRef paTypes As Any, _
                    ByRef paValues As Any, _
                    ByRef fuReturn As Variant)

' ----==== Variablen ====----
Private ole_typ(10) As Integer
Private ole_ptr(10) As Long

' ----==== Init Interface ====----
Public Function InitInterface(ByRef Interface As Interface_Data, ByVal cid As _
    String, ByVal IID As String) As Boolean

    ' Erstellen eines Interfaces aus IID und CLSID

    Dim car() As Byte
    Dim iar() As Byte

    ' Falls die CLSID nicht konvertiert werden konnte
    If Not oleConvert(cid, car()) Then

        ' Fehler auslösen
        Call InterfaceError(Interface, ecd_OleConvert)

    ' Falls die IID nicht konvertiert werden konnte
    ElseIf Not oleConvert(IID, iar()) Then

        ' Fehler auslösen
        Call InterfaceError(Interface, ecd_OleConvert)

    ' Falls das Interface nicht aus CLSID und IID erstellt werden konnte
    ElseIf CoCreateInstance(car(0), 0&, CLSCTX_INPROC, iar(0), Interface.ifc) <> _
        S_OK Then

        ' Fehler auslösen
        Call InterfaceError(Interface, ecd_InitInterface)

    Else

        ' Erstellung des Interfaces war erfolgreich
        InitInterface = True

    End If

End Function

' ----==== Release Interface ====----
Public Function ReleaseInterface(ByRef Interface As Interface_Data)

    Dim lRet As Long

    ' ist ein Pointer auf ein Interfac vorhanden
    If Interface.ifc Then

        ' Funktion Release des Interfaces aufrufen
        Call DispCallFunc(Interface.ifc, IID_Release, CC_STDCALL, vbLong, 0&, 0&, _
            0&, lRet)

    End If

End Function

' ----==== Interface Error ====----
Public Function InterfaceError(ByRef Interface As Interface_Data, Optional ByVal _
    ecd As Interface_errCodes = -1) As Boolean

    Dim dmy As Object
    Dim obj As Object

    With Interface

        ' ist eine Fehlernummer vorhanden
        If ecd Then .ecd = ecd ' Fehlernummer speichern

        ' Felertext nach Fehlernummer speichern
        Select Case .ecd
        
        Case Is < 0:                .etx = "": .ecd = ecd_None
        Case ecd_InvalidCall:       .etx = "invalid function call"
        Case ecd_OleConvert:        .etx = "could not convert classid"
        Case ecd_InitInterface:     .etx = "could not convert interface id"
        Case ecd_OleInvoke:         .etx = "could not invoke ifc function"

        End Select

        If .ecd = ecd_None Then ' nur wenn Fehler ecd_None

        ElseIf Not .RaiseErrors Then ' Nur wenn RaiseErros = False

        ElseIf .owner Then ' ist ein Pointer zu einer Klasse vorhanden
            
            ' Objekt vom Pointer
            Call RtlMoveMemory(dmy, .owner, 4)

            ' Objekt speichern
            Set obj = dmy

            ' Objekt löschen
            Call RtlMoveMemory(dmy, 0&, 4)
            
            ' Sub x_RaiseError in der entsprechenden Klasse aufrufen
            obj.x_RaiseError

        End If

    End With

End Function

' ----==== IID/CLSID to ByteArray ====----
Private Function oleConvert(ByVal cid As String, ByRef bar() As Byte) As Boolean

    ReDim bar(15)
    oleConvert = (CLSIDFromString(StrPtr(cid), bar(0)) = S_OK)

End Function

' ----==== Call Interface Function ====----
Public Function oleInvoke(ByRef Interface As Interface_Data, ByVal cmd As Long, _
    ByRef ret As Variant, ByVal chk As Boolean, ParamArray arr()) As Boolean

    Dim lpc As Long
    Dim var As Variant

    ' wenn kein Interface-Pointer vorhanden ist
    If Interface.ifc = 0 Then

        ' Fehler auslösen
        Call InterfaceError(Interface, ecd_InvalidCall)

    Else

        ' nur wenn zum Aufruf der Interface-Funktion auch
        ' Parameter vorhanden sind.
        If UBound(arr) >= 0 Then

            ' ParamArray nach Variant
            var = arr

            ' ist der Variant ein Array
            If IsArray(var) Then var = var(0)

            ' alle Parameter durchlaufen
            For lpc = 0 To UBound(var)

                ole_typ(lpc) = VarType(var(lpc))    ' Typ des Parameter
                ole_ptr(lpc) = VarPtr(var(lpc))     ' Pointer auf den Parameter
            
            Next

        End If

        ' Funktion des Interfaces aufrufen
        Call DispCallFunc(Interface.ifc, cmd * 4, CC_STDCALL, VarType(ret), lpc, _
            ole_typ(0), ole_ptr(0), ret)

        oleInvoke = True

        If Not chk Then ' wenn chk = False

        ElseIf VarType(ret) <> vbLong Then ' wenn ret <> vbLong ist

        ElseIf ret <> S_OK Then ' wenn ret <> S_OK ist

            ' Fehler auslösen
            Call InterfaceError(Interface, ecd_OleInvoke)

            ' zurück geben das der Aufruf fehlgeschlagen ist
            oleInvoke = False

        End If
                
    End If

End Function

