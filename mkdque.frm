VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mkdque 
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Solutions"
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
      Height          =   9255
      Left            =   4440
      TabIndex        =   3
      Top             =   530
      Width           =   15855
      Begin MSComctlLib.ListView ListView2 
         Height          =   7815
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   13785
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qno."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Question"
            Object.Width           =   24721
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Answer"
            Object.Width           =   13758
         EndProperty
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
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
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
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
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   3015
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
      Caption         =   "Test"
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4215
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
      Caption         =   "Marked Question"
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
Attribute VB_Name = "mkdque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
'On Error GoTo l:
Dim crs As String
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
   ' On Error GoTo m:
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
Dim ss As New ADODB.Recordset
Set ss = Nothing
ListView2.ListItems.Clear
sql = "select * from test where testid = " & ListView1.SelectedItem
Set rs = Nothing
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Text8.Text = rs!testid
Text9.Text = rs!course
Text10.Text = rs!tdate
Text11.Text = rs!noq
Set rs = Nothing
sql = "select * from mque where testid = " & ListView1.SelectedItem & " and uid = " & uid
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   sql = "select * from question where testid = " & ListView1.SelectedItem & " and qno = " & rs![qno]
   ss.Open sql, cn
   Set it = ListView2.ListItems.Add(, , rs![qno])
   it.SubItems(1) = ss![que]
   If ss![ans] = 1 Then
        it.SubItems(2) = ss![op1]
   ElseIf ss![ans] = 2 Then
        it.SubItems(2) = ss![op2]
   ElseIf ss![ans] = 3 Then
        it.SubItems(2) = ss![op3]
   ElseIf ss![ans] = 4 Then
        it.SubItems(2) = ss![op4]
   Else
         'nothing
   End If
   Set ss = Nothing
   rs.MoveNext
Wend
End Sub
