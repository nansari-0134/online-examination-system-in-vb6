VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form result 
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
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   10080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RANK LIST"
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
      TabIndex        =   4
      Top             =   600
      Width           =   15855
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000003&
         Caption         =   "Delete"
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
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8400
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000003&
         Caption         =   "Print"
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
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8400
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   7815
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   13785
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
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Id "
            Object.Width           =   6376
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rank"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Student Name"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Student Mark"
            Object.Width           =   8820
         EndProperty
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
      Left            =   1080
      TabIndex        =   3
      Top             =   9960
      Width           =   1935
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
         TabIndex        =   8
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
      Caption         =   "Test Results"
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
Attribute VB_Name = "result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim ec As New ADODB.Recordset
Dim ec1 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim sql As String
Dim i As Integer
Private Sub Command1_Click()
Unload Me
home.Show
End Sub

Private Sub Command2_Click()
Set rs = Nothing
For i = ListView2.ListItems.Count To 1 Step -1
    If ListView2.ListItems(i).Checked = True Then
        sql = "Delete from test where testid = " & ListView1.SelectedItem
        rs.Open sql, cn
        Set rs = Nothing
    End If
Next i
Set rs = Nothing
For i = ListView2.ListItems.Count To 1 Step -1
    If ListView2.ListItems(i).Checked = True Then
        sql = "Delete from result where testid = " & ListView1.SelectedItem
        rs.Open sql, cn
        Set rs = Nothing
    End If
Next i
Unload Me
result.Show
End Sub

Private Sub Command3_Click()
Set rs = Nothing
i = 0
con.Open "Provider=MSDataShape;Data Provider=NONE;"
sql = "SHAPE APPEND NEW adDouble As testid,NEW adDouble As uid,NEW adInteger As mark,NEW adchar(40) as name,NEW adInteger As rank, " _
      & "((SHAPE APPEND NEW adDouble as uid) As child Relate uid to uid) As data"
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
    Set ec = Nothing
    sql = "Select * from result where testid = " & ListView1.SelectedItem & " order by mark DESC"
    ec.Open sql, cn, adOpenDynamic, adLockOptimistic
    While Not ec.EOF
       sql = "select * from user where userid = " & ec!uid
       ec1.Open sql, cn
       .AddNew
         !testid = ec!testid
         !uid = ec!uid
         !mark = ec!mark
         !Name = ec1![Name]
         !rank = i + 1
       .Update
       i = i + 1
       Set ec1 = Nothing
       ec.MoveNext
    Wend
    ec.MoveFirst
    Set rs1 = rs.Fields("data").Value
End With
     With rs1
        While Not ec.EOF
           .AddNew
           !uid = ec!uid
           .Update
           ec.MoveNext
        Wend
     End With
Set ec = Nothing
Set ec1 = Nothing
sql = "select * from test where testid = " & ListView1.SelectedItem
ec.Open sql, cn, adOpenDynamic, adLockOptimistic
res.Sections("section4").Controls("label9").Caption = ec!course
res.Sections("section4").Controls("label11").Caption = ec!topic
Set ec = Nothing
res.DataMember = rs.DataMember
Set res.DataSource = rs
'res.Sections("section6").Controls("text1").Text.DataMember = rs.DataMember
'res.Sections("section6").Controls("text1").Text.DataField = "uid"
'res.Sections("section6").Controls("text3").Text.DataMember = rs.DataMember
'res.Sections("section6").Controls("text3").Text.DataField = "name"
'res.Sections("section6").Controls("text4").Text.DataMember = rs.DataMember
'res.Sections("section6").Controls("text4").Text.DataField = "mark"
Unload Me
res.Show
End Sub

Private Sub Command4_Click()
Set rs = Nothing
For i = ListView2.ListItems.Count To 1 Step -1
    If ListView2.ListItems(i).Checked = True Then
        sql = "Delete from result where testid = " & ListView1.SelectedItem & " and uid = " & ListView2.ListItems(i)
        rs.Open sql, cn
        Set rs = Nothing
    End If
Next i
Unload Me
result.Show
End Sub

Private Sub Command5_Click()
Set rs = Nothing
Set rs1 = Nothing
Set ec = Nothing
Set ec1 = Nothing
con.Close
End Sub

Private Sub Form_Load()
Dim s As New ADODB.Recordset
Set rs = Nothing
If ut = 0 Then
sql = "select * from user where userid = " & uid
s.Open sql, cn, adOpenDynamic, adLockOptimistic
sql = "select * from test where course = '" & s!course & "'"
Set s = Nothing
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set it = ListView1.ListItems.Add(, , rs!testid)
   it.SubItems(1) = rs!topic
   it.SubItems(2) = rs!tdate
   rs.MoveNext
Wend
Set rs = Nothing
Else
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
If ut = 1 Then
   Command2.Visible = True
   Command4.Visible = True
Else
   Command2.Visible = False
   Command4.Visible = False
End If
End Sub
Private Sub ListView1_Click()
Dim s As New ADODB.Recordset
Dim i As Integer
ListView2.ListItems.Clear
i = 1
sql = "select * from result where testid = " & ListView1.SelectedItem & " order by mark DESC"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Set it = ListView2.ListItems.Add(, , rs!uid)
    sql = "select * from user where userid = " & rs!uid
    s.Open sql, cn
    it.SubItems(1) = i
    it.SubItems(2) = s![Name]
    it.SubItems(3) = rs!mark
    Set s = Nothing
    i = i + 1
    rs.MoveNext
Wend
Set rs = Nothing
ListView2.SetFocus
For i = 1 To ListView2.ListItems.Count
 If ListView2.ListItems(i) = uid Then
     Set ListView2.SelectedItem = ListView2.ListItems(i)
     Exit Sub
 End If
Next i
End Sub
