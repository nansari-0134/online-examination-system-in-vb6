VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tests 
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   10080
      TabIndex        =   25
      Top             =   10200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   10680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7920
      Top             =   10560
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Test Details"
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
      Height          =   9135
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   15855
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2160
         Width           =   15495
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8710
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   8760
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
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
         TabIndex        =   21
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "hrs."
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
         Left            =   10080
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "mins."
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
         Left            =   11280
         TabIndex        =   19
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Details :"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Topic :"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Questions :"
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
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Course:"
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
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Id :"
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
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   9135
         Left            =   0
         TabIndex        =   23
         Top             =   0
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
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Test Id"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Topic"
            Object.Width           =   4074
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1949
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Tests"
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
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
Attribute VB_Name = "tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim t As Boolean
Dim b As Integer
Private Sub Command1_Click()
Unload Me
home.Show
End Sub

Private Sub Command3_Click()
tid = CLng(Text8.Text)
sql = "select * from question where testid = " & Text8.Text & " order by qno"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
cnt = 0
'time = rs!ttime * 60
Call fun
End Sub
Private Sub fun()
If rs.EOF = False Then
   qno = rs!qno
   que.Show
   rs.MoveNext
   t = True
   Timer1.Enabled = True
Else
  If cnt = 0 Then
   Unload Me
   tests.Show
   MsgBox "Congratualations! You have completed your test successfully", vbInformation + vbOKOnly
   Set rs = Nothing
   Timer1.Enabled = False
   Call up
  Else
    qno = q(b)
    cnt = cnt - 1
    b = b + 1
    que.Show
    t = True
    Timer1.Enabled = True
  End If
End If
End Sub
Private Sub up()
sql = "insert into result(uid,testid,mark) values(" & uid & "," & tid & "," & tm & ")"
rs.Open sql, cn
Set rs = Nothing
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Set rs = Nothing
Call up
MsgBox "Congratualations! You have completed your test successfully", vbInformation + vbOKOnly
Unload Me
tests.Show
End Sub

Private Sub Command5_Click()
If t = True Then
 t = False
End If
End Sub
Private Sub Form_Load()
On Error GoTo l:
Dim crs As String
b = 0
If ut = 0 Then
    sql = "select * from user where userid = " & uid
    Set rs = Nothing
    rs.Open sql, cn
    crs = rs!course
    Set rs = Nothing
    sql = "select * from test where course = '" & crs & "'"
    rs.Open sql, cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
       Set it = ListView1.ListItems.Add(, , rs!testid)
       it.SubItems(1) = rs!topic
       it.SubItems(2) = rs!tdate
       rs.MoveNext
    Wend
    Set rs = Nothing
Else
l:
    On Error GoTo m:
    Set rs = Nothing
    sql = "select * from test"
    rs.Open sql, cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
       Set it = ListView1.ListItems.Add(, , rs!testid)
       it.SubItems(1) = rs!topic
       it.SubItems(2) = rs!tdate
       rs.MoveNext
    Wend
Set rs = Nothing
End If
Exit Sub
m:
End Sub
Private Sub ListView1_Click()
sql = "select * from test where testid = " & ListView1.SelectedItem
Set rs = Nothing
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Text8.Text = rs!testid
Text9.Text = rs!course
Text10.Text = rs!tdate
Text11.Text = rs!noq
Text12.Text = rs!ttime \ 60
Text13.Text = rs!ttime Mod 60
Text15.Text = rs!topic
time = rs!ttime * 60
Set rs = Nothing
sql = "select sum(mark) from question where testid = " & ListView1.SelectedItem
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
    Label7.Caption = "Total Marks : " & rs.Fields(0)
Set rs = Nothing
sql = "select * from result where testid = " & Text8.Text & " and uid = " & uid
rs.Open sql, cn
If rs.EOF = True Then
    Command3.Enabled = True
Else
    Command3.Enabled = False
End If
Set rs = Nothing
End Sub
Private Sub Timer1_Timer()
If t = False Then
   Call fun
End If
End Sub
