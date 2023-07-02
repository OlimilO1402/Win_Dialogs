Attribute VB_Name = "MApp"
Option Explicit
Public Const FileExtFilter As String = "Textfile (*.txt)|*.txt|html-file (*.htm, *.html)|*.htm*|All files (*.*)|*.*"

Sub Main()
    FMain.Show
End Sub

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function TaskDialog(Title As String, Instruction As String, Content As String, Optional ByVal Icon As ETaskDialogIcon, Optional ByVal Buttons As ETaskDialogButton) As TaskDialogSE
    Set TaskDialog = New TaskDialogSE: TaskDialog.New_ Title, Instruction, Content, Icon, Buttons
End Function

