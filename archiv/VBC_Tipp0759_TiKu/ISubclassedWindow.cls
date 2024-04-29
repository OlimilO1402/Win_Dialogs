VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclassedWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function HandleMessage(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal scID As ESubclassID, ByRef bCallDefProc As Boolean) As Long: End Function
