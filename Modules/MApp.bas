Attribute VB_Name = "MApp"
Option Explicit
Public Const FileExtFilter As String = "Textfile (*.txt)|*.txt|html-file (*.htm, *.html)|*.htm*|All files (*.*)|*.*"


Sub Main()
    Form1.Show
End Sub

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
