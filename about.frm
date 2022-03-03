VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   18915
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   3480
      Picture         =   "about.frx":0000
      Top             =   1200
      Width           =   12750
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
home.show
End Sub
