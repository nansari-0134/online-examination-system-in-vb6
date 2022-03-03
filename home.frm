VERSION 5.00
Begin VB.Form home 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleMode       =   0  'User
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   10000
      Width           =   22335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9240
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Manage tests"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   14040
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Stuents"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   16200
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   18600
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Online Examination System"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " Feedback"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   6
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Notification"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   5
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " Edit Profile "
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Marked Quetions"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   3
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Result "
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " Tests "
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   6840
      Left            =   0
      Picture         =   "home.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20610
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Private Sub Form_Load()
'Me.WindowState = 2
Me.Width = Screen.Width - 80
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 400
Me.Left = -100
Me.Top = -100
sql = "select * from user where userid = " & uid
rs.Open sql, cn
Label13.Caption = "  Current User :- " & rs![Name]
Set rs = Nothing
If ut = 1 Then
   Label10.Visible = True
   Label11.Visible = True
Else
   Label10.Visible = False
   Label11.Visible = False
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFFFFFF
Label1.BackStyle = 0
Label2.BackStyle = 0
Label4.BackStyle = 0
Label5.BackStyle = 0
Label6.BackStyle = 0
Label7.BackStyle = 0
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFFFFFF
Label10.ForeColor = &HFFFFFF
Label11.ForeColor = &HFFFFFF
Label1.BackStyle = 0
Label2.BackStyle = 0
Label4.BackStyle = 0
Label5.BackStyle = 0
Label6.BackStyle = 0
Label7.BackStyle = 0
End Sub
Private Sub Label1_Click()
Unload Me
tests.Show
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HE0E0E0
Label1.BackStyle = 1
End Sub
Private Sub Label10_Click()
mngstd.Show
End Sub
Private Sub Label11_Click()
addque.Show
Unload home
End Sub
Private Sub Label12_Click()
Unload Me
about.Show
End Sub
Private Sub Label2_Click()
Unload Me
result.Show
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HE0E0E0
Label2.BackStyle = 1
End Sub
Private Sub Label4_Click()
Unload Me
mkdque.Show
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = &HE0E0E0
Label4.BackStyle = 1
End Sub
Private Sub Label5_Click()
Unload Me
edtprf.Show
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &HE0E0E0
Label5.BackStyle = 1
End Sub
Private Sub Label6_Click()
Unload Me
mngnoti.Show
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = &HE0E0E0
Label6.BackStyle = 1
End Sub
Private Sub Label7_Click()
Unload Me
feedback.Show
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &HE0E0E0
Label7.BackStyle = 1
End Sub
Private Sub Label9_Click()
Unload Me
Unload mdi
login.Show
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &H404040
Label10.ForeColor = &HFFFFFF
Label11.ForeColor = &HFFFFFF
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H404040
Label9.ForeColor = &HFFFFFF
Label11.ForeColor = &HFFFFFF
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &H404040
Label9.ForeColor = &HFFFFFF
Label10.ForeColor = &HFFFFFF
End Sub
