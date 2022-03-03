VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mngstd 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      Caption         =   "Print Student Info"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1920
      TabIndex        =   37
      Top             =   4080
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command9 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   41
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   40
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         TabIndex        =   39
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Course :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Student Info List"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   36
      Top             =   9960
      Width           =   3015
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9135
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   16113
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Id"
         Object.Width           =   1859
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5574
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refute"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   4440
      TabIndex        =   8
      Top             =   5040
      Width           =   15855
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   9360
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000004&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14280
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   31
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2880
         Width           =   13215
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9360
         TabIndex        =   27
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9360
         TabIndex        =   18
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000004&
         Caption         =   "update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14280
         TabIndex        =   11
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Verification :"
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
         Left            =   7800
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer :"
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
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Security Question :"
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
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
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
         Left            =   7800
         TabIndex        =   28
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Id :"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name :"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Id :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No :"
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
         Left            =   7800
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Course :"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   4440
      TabIndex        =   4
      Top             =   530
      Width           =   15855
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student I d "
            Object.Width           =   5509
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student name "
            Object.Width           =   9641
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Course"
            Object.Width           =   5509
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Mobile"
            Object.Width           =   6879
         EndProperty
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9360
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000004&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
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
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No :"
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
         Left            =   7680
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name :"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove  "
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "<-  Back  "
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Students"
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
      TabIndex        =   2
      Top             =   0
      Width           =   8295
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   21000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   21000
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   480
      Y2              =   11400
   End
End
Attribute VB_Name = "mngstd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim a As PrinterControl
Private Sub Command1_Click()
Unload Me
home.Show
End Sub
Private Sub Command2_Click()
Dim a As Integer
Set rs = Nothing
a = MsgBox("Are You sure ?" & vbNewLine & "You want to delete record(s)", vbCritical + vbOKCancel, "Attention!")
If a = vbOK Then
   For i = ListView1.ListItems.Count To 1 Step -1
       If ListView1.ListItems.Item(i).Checked = True Then
           sql = "delete from user where userid = " & ListView1.ListItems(i)
           rs.Open sql, cn
           Set rs = Nothing
       End If
   Next i
End If
Call data
Call show1
End Sub
Private Sub Command4_Click()
Dim i As Integer
Set rs = Nothing
Set a = New PrinterControl
i = InputBox("1.Print All Student Info" & vbNewLine & "2. Print Info of students of specific course", "Print")
If i = 1 Then
   'print all
   sql = "select * from user where course <> 'admin' order by userid"
   rs.Open sql, cn, adOpenDynamic, adLockOptimistic
   a.ChngOrientationPortrait
   stud.DataMember = rs.DataMember
   Set stud.DataSource = rs
   Unload Me
   stud.Show
ElseIf i = 2 Then
   'specified course
   Frame4.Visible = True
   Frame4.Top = Screen.Height / 2 - Frame4.Height / 2
   Frame4.Left = Screen.Width / 2 - Frame4.Width / 2
   sql = "select distinct course from user where course <> 'admin'"
   rs.Open sql, cn, adOpenDynamic, adLockOptimistic
   Combo1.Clear
   While Not rs.EOF
       Combo1.AddItem rs.Fields(0)
       rs.MoveNext
   Wend
End If
Set rs = Nothing
End Sub

Private Sub Command6_Click()
Dim a As Integer
Set rs = Nothing
a = MsgBox("Are You sure ?" & vbNewLine & "You want to Refute record(s)", vbCritical + vbOKCancel, "Attention!")
If a = vbOK Then
   For i = ListView1.ListItems.Count To 1 Step -1
       If ListView1.ListItems.Item(i).Checked = True Then
           sql = "update user set verification = 0 where userid = " & ListView1.ListItems(i)
           rs.Open sql, cn
           Set rs = Nothing
       End If
   Next i
End If
End Sub
Private Sub Command7_Click()
Set rs = Nothing
Dim a As Integer
a = MsgBox("Do You Want To update Record !", vbInformation + vbOKCancel, "Alert")
If a = vbOK Then
sql = "update user set name = '" & Text5.Text & "'" _
       & ", course = '" & Text2.Text & "'" _
       & ", email = '" & Text6.Text & "'" _
       & ", course = '" & Text2.Text & "'" _
       & ", qans = '" & Text11.Text & "'" _
       & ", verification = " & Check1.Value & "" _
       & ", mobile = '" & Text3.Text & "'" _
       & ", password = '" & Text7.Text & "' where userid = " & CLng(Text4.Text)
rs.Open sql, cn
Set rs = Nothing
End If
End Sub

Private Sub Command8_Click()
If Combo1.Text <> "" Then
 sql = "select * from user where course = '" & Combo1.Text & "' order by userid"
   rs.Open sql, cn, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
        stud.DataMember = rs.DataMember
        Set stud.DataSource = rs
        Unload Me
        stud.Show
   Else
         MsgBox "No Student Enrolled to the course", vbInformation + vbOKOnly, "Print"
   End If
Else
   MsgBox "No Course Selected", vbInformation + vbOKOnly, "Print"
End If
End Sub

Private Sub Command9_Click()
Frame4.Visible = False
Combo1.Text = ""
End Sub

Private Sub Form_Load()
Set rs = Nothing
Call data
Call show1
End Sub
Private Sub data()
ListView1.ListItems.Clear
ListView2.ListItems.Clear
sql = "select * from user where course <> 'admin'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set Item = ListView1.ListItems.Add(, , rs!userid)
   Item.SubItems(1) = rs![Name]
   Set it = ListView2.ListItems.Add(, , rs!userid)
   it.SubItems(1) = rs!Name
   it.SubItems(2) = rs!course
   it.SubItems(3) = rs!mobile
   rs.MoveNext
Wend
Set rs = Nothing
End Sub
Private Sub show1()
'on error goto l:
sql = "select * from user where userid = " & ListView1.SelectedItem
rs.Open sql, cn
Text4.Text = rs![userid]
Text5.Text = rs![Name]
Text2.Text = rs![course]
Text6.Text = rs![email]
Text9.Text = rs![qid]
Text11.Text = rs![qans]
Text3.Text = rs![mobile]
Text7.Text = rs![password]
Check1.Value = rs!verification
Set rs = Nothing
Exit Sub
l:
End Sub
Private Sub ListView1_Click()
Call show1
End Sub
Private Sub Text1_Change()
On Error GoTo l:
Set rs = Nothing
ListView2.ListItems.Clear
sql = "select * from user where course <> 'admin' and course like '" & Text1.Text & "%'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set it = ListView2.ListItems.Add(, , rs!userid)
   it.SubItems(1) = rs!Name
   it.SubItems(2) = rs!course
   it.SubItems(3) = rs!mobile
   rs.MoveNext
Wend
Set rs = Nothing
Exit Sub
l:
End Sub
Private Sub Text10_Change()
On Error GoTo l:
Set rs = Nothing
ListView2.ListItems.Clear
sql = "select * from user where course <> 'admin' and mobile like '" & Text10.Text & "%'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set it = ListView2.ListItems.Add(, , rs!userid)
   it.SubItems(1) = rs!Name
   it.SubItems(2) = rs!course
   it.SubItems(3) = rs!mobile
   rs.MoveNext
Wend
Set rs = Nothing
Exit Sub
l:
End Sub

Private Sub text10_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text8_Change()
On Error GoTo l:
Set rs = Nothing
ListView2.ListItems.Clear
sql = "select * from user where course <> 'admin' and name like '" & Text8.Text & "%'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set it = ListView2.ListItems.Add(, , rs!userid)
   it.SubItems(1) = rs!Name
   it.SubItems(2) = rs!course
   it.SubItems(3) = rs!mobile
   rs.MoveNext
Wend
Set rs = Nothing
Exit Sub
l:
End Sub
Private Function validate() As Boolean
If Text5.Text = "" Then
   MsgBox "Username is missing", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text7.Text = "" Then
   MsgBox "Password is missing", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text11.Text = "" Then
   MsgBox "Security Answer is missing", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
validate = True
End Function
