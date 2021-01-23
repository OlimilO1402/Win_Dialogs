Attribute VB_Name = "MObjPtr"
Option Explicit

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef src As Any, ByVal bytLength As Long)

Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal bytLength As Long)

Public Function PtrToObject(ByVal p As Long) As Object
    RtlMoveMemory ByVal VarPtr(PtrToObject), p, 4
End Function

Public Sub ZeroToObject(obj As Object) 'As Object
    RtlZeroMemory ByVal VarPtr(obj), 4
End Sub
