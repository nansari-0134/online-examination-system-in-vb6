VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form feedback 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   18930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Ubuntu"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   18930
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Submit"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<- Back"
      Height          =   495
      Left            =   18960
      TabIndex        =   12
      Top             =   9600
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   9480
      TabIndex        =   11
      Top             =   5280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User id "
         Object.Width           =   5951
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Feedback"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   1935
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   9135
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excellent"
         Height          =   615
         Left            =   7440
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Very Good"
         Height          =   615
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Good"
         Height          =   615
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fine"
         Height          =   615
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bad"
         Height          =   615
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Very Bad"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comment"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   6480
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Give Rating"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "What is your opinion  :-"
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
      TabIndex        =   14
      Top             =   5880
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "We would like your Feedback to improve Our Quality..."
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
      TabIndex        =   13
      Top             =   5280
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "feedback.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim sql As String
Private Sub Command1_Click()
Set rs = Nothing
If Option1.Value = True Then
   If Option3.Value = True Then
        sql = "insert into feedback(userid,content) values(" & uid & ",'" & Option3.Caption & "')"
        rs.Open sql, cn
        Set rs = Nothing
        MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
   ElseIf Option4.Value = True Then
        sql = "insert into feedback(userid,content) values(" & uid & ",'" & Option4.Caption & "')"
        rs.Open sql, cn
        Set rs = Nothing
        MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
   ElseIf Option5.Value = True Then
        sql = "insert into feedback(userid,content) values(" & uid & ",'" & Option5.Caption & "')"
        rs.Open sql, cn
        Set rs = Nothing
        MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
   ElseIf Option6.Value = True Then
        sql = "insert into feedback(userid,content) values(" & uid & ",'" & Option6.Caption & "')"
        rs.Open sql, cn
        Set rs = Nothing
        MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
   ElseIf Option7.Value = True Then
        sql = "insert into feedback(userid,content) values(" & uid & ",'" & Option7.Caption & "')"
        rs.Open sql, cn
        Set rs = Nothing
        MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
   End If
Else
 If Text1.Text <> "" Then
   sql = "insert into feedback(userid,content) values(" & uid & ",'" & Text1.Text & "')"
   rs.Open sql, cn
   Set rs = Nothing
   MsgBox "Your Feedback is sumbitted succefully", vbInformation + vbOKOnly, "Success"
 Else
   MsgBox "Please Type something first", vbInformation + vbOKOnly, "Alert"
 End If
End If
Call data
End Sub
Private Sub Command3_Click()
Unload Me
home.Show
End Sub
Private Sub Form_Load()
        If ut = 1 Then
           ListView1.Visible = True
        Else
           ListView1.Visible = False
        End If
Call data
End Sub
Private Sub data()
Dim it As ListItem
 Set rs = Nothing
 ListView1.ListItems.Clear
 sql = "select * from feedback"
 rs.Open sql, cn, adOpenDynamic, adLockOptimistic
 While Not rs.EOF
   Set it = ListView1.ListItems.Add(, , rs!userid)
    it.SubItems(1) = rs!content
    rs.MoveNext
 Wend
Set rs = Nothing
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
Frame1.Visible = True
Frame2.Visible = False
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
Frame1.Visible = False
Frame2.Visible = True
End If
End Sub
