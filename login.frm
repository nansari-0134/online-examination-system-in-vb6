VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Ubuntu"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "login.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   360
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox Text12 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   480
         TabIndex        =   37
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Request For Password"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2470
         Width           =   6855
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   480
         TabIndex        =   35
         Top             =   1920
         Width           =   6615
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "login.frx":2B15C
         Left            =   360
         List            =   "login.frx":2B1A2
         TabIndex        =   34
         Text            =   "Choose Security Quetion For Password Reset."
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label Label19 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7350
         TabIndex        =   39
         Top             =   20
         Width           =   200
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   120
         Width           =   1335
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4335
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Shape Shape10 
         Height          =   3015
         Left            =   0
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   8520
      TabIndex        =   19
      Top             =   600
      Width           =   7095
      Begin MSComCtl2.DTPicker text10 
         Height          =   495
         Left            =   3600
         TabIndex        =   11
         Top             =   3480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   115474433
         CurrentDate     =   43763
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Sign Up"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7200
         Width           =   6855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   240
         TabIndex        =   14
         Top             =   6480
         Width           =   6615
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "login.frx":2B50C
         Left            =   120
         List            =   "login.frx":2B552
         TabIndex        =   13
         Text            =   "Choose Security Quetion For Password Reset."
         Top             =   5880
         Width           =   6855
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   360
         TabIndex        =   12
         Top             =   4680
         Width           =   6375
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   3720
         TabIndex        =   9
         Top             =   2475
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   3530
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   3720
         TabIndex        =   7
         Top             =   1515
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2495
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DoB :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   32
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   6855
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   6855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   2445
         Width           =   2655
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   30
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   1485
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mobile No."
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   29
         Top             =   1220
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Confirm Password :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2445
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1485
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1220
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
         Caption         =   "        Be a part of your institution"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Caption         =   "    Sign Up"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   615
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   1215
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Frame lg 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "login.frx":2B8BC
         Left            =   120
         List            =   "login.frx":2B8C6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4320
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   160
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3720
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Forget Password ?"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   495
         Left            =   2760
         MouseIcon       =   "login.frx":2B8DB
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log in As :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1420
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "        See Whats Going on in your institution"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "    Login"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   615
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   1215
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   4575
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim qid As Long
Private Sub Command1_Click()
uid = CLng(Text2.Text)
If Combo1.Text = "Student" Then
    Set rs = Nothing
    rs.Open "Select * from user where userid = " & Text2.Text & " and password = '" & Text3.Text & "' and course <> 'Admin'", cn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then '
      If rs!verification = 1 Then
        home.Label10.Visible = False
        home.Label11.Visible = False
        ut = 0
        Unload Me
        mdi.Show
      Else
        MsgBox "You are not verified yet." & vbNewLine & "Please contact to admin", vbInformation + vbOKOnly, "Info"
      End If
    Else
     MsgBox "Wrong Username And Password", vbCritical + vbOKOnly
     Set rs = Nothing
    End If
ElseIf Combo1.Text = "Admin" Then
    Set rs = Nothing
    rs.Open "Select * from user where userid = " & Text2.Text & " and password = '" & Text3.Text & "' and course = 'Admin'", cn, adOpenDynamic, adLockOptimistic
  If rs.EOF = False Then '
    home.Label10.Visible = True
    home.Label11.Visible = True
    ut = 1
    Unload Me
    mdi.Show
  Else
     MsgBox "Wrong Username And Password", vbCritical + vbOKOnly
     Set rs = Nothing
  End If
Else
    Set rs = Nothing
    rs.Open "Select * from user where userid = " & Text2.Text & " and password = '" & Text3.Text & "' and course <> 'Admin'", cn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then '
      If rs!verification = 1 Then
        home.Label10.Visible = False
        home.Label11.Visible = False
        ut = 0
        Unload Me
        mdi.Show
      Else
        MsgBox "You are not verified yet." & vbNewLine & "Please contact to admin", vbInformation + vbOKOnly, "Info"
      End If
    Else
     MsgBox "Wrong Username And Password", vbCritical + vbOKOnly
     Set rs = Nothing
    End If
End If
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H80000015
End Sub
Private Sub Command2_Click()
Set rs = Nothing
If validate = True Then
'call validate
            sql = "select * from verify where question= '" & Combo2.Text & "'"
            rs.Open sql, cn
            qid = rs!qid
            Set rs = Nothing
            sql = "select max(userid) from user"
            rs.Open sql, cn
            uid = rs.Fields(0)
            Set rs = Nothing
            uid = uid + 1
            sql = "insert into `user`(`userid`,`name`,`mobile`,`email`,`dob`,`address`,`qid`,`qans`,`verification`,`course`,`password`)" _
                   & " values (" & uid & ",'" & Text1.Text & "','" & Text5.Text _
                   & "','" & Text7.Text & "','" & Text10.Value & "','" _
                   & Text8.Text & "'," _
                   & qid & ",'" & Text9.Text & "',0,'','" & Text6.Text & "')"
            rs.Open sql, cn
            MsgBox "Sign Up Successful" & vbNewLine & "" & vbNewLine & "Wait for verification. After Verification you can login with your Id and Password"
            Text1.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Combo2.Text = "Choose Security Quetion For Password Reset."
End If
End Sub
Private Sub Command3_Click()
Dim up As New ADODB.Recordset
Dim n As Integer
sql = "select * from verify where question = '" & Combo3.Text & "'"
up.Open sql, cn
If up.EOF = True Then
   MsgBox "Something went wrong", vbInformation + vbOKOnly, "Error"
   Exit Sub
End If
n = up!qid
Set up = Nothing
sql = "select * from user where userid = " & Text12.Text & " and qid = " & n
up.Open sql, cn
If up.EOF = True Then
   MsgBox "Wrong Inputs", vbCritical + vbOKOnly, "Error"
   Set up = Nothing
   Exit Sub
End If
If up!qans = Text11.Text Then
   MsgBox "Your Password Is : " & up!password
Else
   MsgBox "Wrong Answer", vbCritical + vbOKOnly, "Wrong"
End If
End Sub
Private Sub Form_Load()
Frame1.Left = lg.Width + 7800
Combo1.Text = Combo1.List(0)
Me.Visible = True
Text2.SetFocus
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H80000015
Label19.BackColor = &H8000000F
End Sub
Private Sub Form_Unload(Cancel As Integer)
'
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H80000015
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.BackColor = &H8000000F
End Sub

Private Sub Label10_Click()
Frame2.Visible = True
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H80000012
End Sub
Private Sub Label19_Click()
Frame2.Visible = False
Combo3.Text = Clear
Combo3.Text = "Choose Security Quetion For Password Reset."
Text12.Text = Clear
Text11.Text = Clear
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.BackColor = &HFFFFFF
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then
   Text3.SetFocus
ElseIf (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Command1.Value = True
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Len(Text5.Text) <= 9 Then
    If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
       KeyAscii = KeyAscii
    Else
       KeyAscii = 0
    End If
Else
    If (KeyAscii = 8) Then
       KeyAscii = KeyAscii
    Else
       KeyAscii = 0
    End If
End If
End Sub
Private Function validate() As Boolean
If Text4.Text = "" Then
   MsgBox "Password is neccessary.." & vbNewLine & "Please enter password" & vbNewLine & "Minimum : 6 character", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text4.Text <> Text6.Text Then
   MsgBox "Password Doesn't Same", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Not Len(Text4.Text) > 5 Then
   MsgBox "Password should contain 6 characters", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text1.Text = "" Then
    MsgBox "Enter User Name Please", vbCritical + vbOKOnly, "Error"
    validate = False
   Exit Function
End If
If Combo2.Text = "Choose Security Quetion For Password Reset." Then
    MsgBox "Please Choose security question", vbCritical + vbOKOnly, "Error"
    validate = False
   Exit Function
End If
If Text9.Text = "" Then
    MsgBox "Please enter security question's answer", vbCritical + vbOKOnly, "Error"
    validate = False
   Exit Function
End If
validate = True
End Function
