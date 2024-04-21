VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ShowFind Me, FR_DOWN Or FR_SHOWHELP, "henderit"
End Sub

Private Sub Command2_Click()
    ShowFind Me, FR_SHOWHELP, "henderit", True, "hendererit"
End Sub

Private Sub Form_Load()
    Caption = "Find/Replace dialogs"
    Command1.Caption = "Find"
    Command2.Caption = "Replace"
    Text1.Text = "Lorem ipsum dolor sit amet, consectetur adipisici elit," & vbCrLf & _
                "sed eiusmod tempor incidunt ut labore et dolore magna" & vbCrLf & _
                "aliqua. Ut enim ad minim veniam, quis nostrud" & vbCrLf & _
                "exercitation ullamco laboris nisi ut aliquid ex ea" & vbCrLf & _
                "commodi consequat. Quis aute iure reprehenderit in" & vbCrLf & _
                "voluptate velit esse cillum dolore eu fugiat nulla" & vbCrLf & _
                "pariatur. Excepteur sint obcaecat cupiditat non" & vbCrLf & _
                "proident, sunt in culpa qui officia deserunt mollit" & vbCrLf & _
                "anim id est laborum." & vbCrLf & _
                "" & vbCrLf & _
                "--" & vbCrLf & _
                "" & vbCrLf & _
                "Duis autem vel eum iriure dolor in henderit in" & vbCrLf & _
                "vulputate velit esse molestie consequat, vel illum" & vbCrLf & _
                "dolore eu feugiat nulla facilisis at vero eros et" & vbCrLf & _
                "accumsan et iusto odio dignissim qui blandit praesent" & vbCrLf & _
                "luptatum zzril delenit augue duis dolore te feugait" & vbCrLf & _
                "nulla facilisi. Lorem ipsum dolor sit amet, consectetuer" & vbCrLf & _
                "adipiscing elit, sed diam nonummy nibh euismod tincidunt" & vbCrLf & _
                "ut laoreet dolore magna aliquam erat volutpat."


'Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit augue duis dolore te feugait nulla facilisi.
'
'Nam liber tempor cum soluta nobis eleifend option congue nihil imperdiet doming id quod mazim placerat facer possim assum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat.
'
'Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis.
'
'At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, At accusam aliquyam diam diam dolore dolores duo eirmod eos erat, et nonumy sed tempor et et invidunt justo labore Stet clita ea et gubergren, kasd magna no rebum. sanctus sea sed takimata ut vero voluptua. est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat.
'
'Consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.
End Sub
