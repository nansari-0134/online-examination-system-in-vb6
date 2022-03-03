VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mngnoti 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
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
   MDIChild        =   -1  'True
   ScaleHeight     =   11070
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   5640
      TabIndex        =   18
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton Command5 
         Caption         =   "<- Back"
         Height          =   615
         Left            =   4440
         TabIndex        =   19
         Top             =   9240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   18000
         Left            =   0
         Picture         =   "mngnoti.frx":0000
         Top             =   -120
         Width           =   14910
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   14400
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command4 
         Caption         =   "<- Back"
         Height          =   615
         Left            =   4440
         TabIndex        =   12
         Top             =   9240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   9240
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   5520
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6376
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Notice Id"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Notice"
            Object.Width           =   4895
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   615
         Left            =   4440
         TabIndex        =   7
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Notification :"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   5160
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Notification :"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   12360
      TabIndex        =   17
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   12360
      TabIndex        =   16
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   12360
      TabIndex        =   15
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   12360
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   12360
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Width           =   14175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   14175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   14175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   14175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "mngnoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As New ADODB.Recordset
Dim s As New ADODB.Recordset
Dim i As Integer

Private Sub Command1_Click()
If Text1.Text <> "" Then
    sql = "select max(nid) from notice"
    s.Open sql, cn
    sql = "insert into notice(nid,note,ndate) values(" & s.Fields(0) + 1 & ",'" & Text1.Text & "','" & Format(Date, "yyyy-mm-dd") & "')"
    rs.Open sql, cn
    Set rs = Nothing
    Set s = Nothing
    MsgBox "Notification updated", vbInformation + vbOKOnly, "Message"
    Unload Me
    mngnoti.Show
Else
    MsgBox "Enter Notice first", vbCritical + vbOKOnly, "Error"
End If
End Sub

Private Sub Command3_Click()
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        sql = "delete from notice where nid = " & ListView1.ListItems(i)
        rs.Open sql, cn
        Set rs = Nothing
    End If
Next i
Unload Me
mngnoti.Show
End Sub
Private Sub Command4_Click()
Unload Me
home.Show
End Sub
Private Sub Command5_Click()
Unload Me
home.Show
End Sub
Private Sub Form_Load()
Me.Width = Screen.Width - 80
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 400
Me.Left = -100
Me.Top = -100
Call data
Call lv
Frame2.Left = Frame1.Left
Frame2.Top = Frame1.Top
If ut = 0 Then
    Frame1.Visible = False
    Frame2.Visible = True
ElseIf ut = 1 Then
    Frame2.Visible = False
    Frame1.Visible = True
End If
End Sub
Private Sub lv()
sql = "select * from notice"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Set it = ListView1.ListItems.Add(, , rs!nid)
    it.SubItems(1) = rs!note
    it.SubItems(2) = rs!ndate
  rs.MoveNext
Wend
Set rs = Nothing
End Sub
Private Sub data()
i = 1
sql = "select * from notice order by nid"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   rs.MoveLast
End If
While Not rs.BOF
     If i = 1 Then
         Label1.Caption = rs!note
         Label8.Caption = rs!ndate
     ElseIf i = 2 Then
         Label2.Caption = rs!note
         Label9.Caption = rs!ndate
     ElseIf i = 3 Then
         Label3.Caption = rs!note
         Label10.Caption = rs!ndate
     ElseIf i = 4 Then
         Label4.Caption = rs!note
         Label11.Caption = rs!ndate
     ElseIf i = 5 Then
         Label5.Caption = rs!note
         Label12.Caption = rs!ndate
     Else
          Set rs = Nothing
          Exit Sub
     End If
     i = i + 1
   rs.MovePrevious
Wend
Set rs = Nothing
End Sub

