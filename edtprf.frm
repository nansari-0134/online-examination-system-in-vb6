VERSION 5.00
Begin VB.Form edtprf 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   20130
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   10920
      TabIndex        =   1
      Top             =   1680
      Width           =   7095
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3720
         TabIndex        =   9
         Top             =   5445
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3720
         TabIndex        =   8
         Top             =   1515
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   5445
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   6
         Top             =   2475
         Width           =   5895
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   240
         TabIndex        =   5
         Top             =   3600
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
         ItemData        =   "edtprf.frx":0000
         Left            =   120
         List            =   "edtprf.frx":0046
         TabIndex        =   4
         Top             =   6240
         Width           =   6855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   3
         Top             =   6840
         Width           =   6615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Update"
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
         TabIndex        =   2
         Top             =   7440
         Width           =   6855
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   6990
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line4 
         X1              =   6960
         X2              =   6960
         Y1              =   4680
         Y2              =   6135
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6990
         Y1              =   6135
         Y2              =   6135
      End
      Begin VB.Line Line2 
         X1              =   6975
         X2              =   120
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   120
         Y1              =   4680
         Y2              =   6150
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Password :"
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
         TabIndex        =   18
         Top             =   5160
         Width           =   2175
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
         TabIndex        =   17
         Top             =   -120
         Width           =   7095
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
         TabIndex        =   16
         Top             =   1220
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New Password :"
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
         TabIndex        =   15
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Change Password :"
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
         TabIndex        =   14
         Top             =   4680
         Width           =   2175
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
         TabIndex        =   13
         Top             =   1220
         Width           =   1455
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
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
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
         TabIndex        =   11
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H8000000A&
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   6855
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Profile ...."
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   -120
      Top             =   0
      Width           =   22815
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   240
      Picture         =   "edtprf.frx":03B0
      Top             =   2880
      Width           =   7500
   End
End
Attribute VB_Name = "edtprf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
'Dim cn As New ADODB.Connection
Dim sql As String
Dim qid As Long
Private Sub Command1_Click()
Unload Me
home.Show
End Sub
Private Sub Command2_Click()
Dim a As Integer
sql = "select * from verify where question= '" & Combo2.Text & "'"
rs.Open sql, cn
qid = rs!qid
Set rs = Nothing
sql = "update user set name = '" & Text1.Text & "' ,mobile = '" & Text5.Text & "' ,email = '" & Text7.Text & "' ,address = '" & Text8.Text & "' ,qid = " & qid & " ,qans = '" & Text9.Text & "' where userid = " & uid
a = MsgBox("Do you want to update Profile ? ", vbExclamation + vbOKCancel, "Confirmation")
If a = vbOK Then
  rs.Open sql, cn
  Set rs = Nothing
Else
   Set rs = Nothing
   Exit Sub
End If
If Text6.Text <> "" Or Text4.Text <> "" Then
  If Text6.Text = "" Then
     MsgBox "If you Want to change password you need to enter previous password", vbCritical + vbOKOnly, "Alert"
     Exit Sub
  ElseIf Text7.Text = "" Then
     MsgBox "If you Want to change password you need to enter new password", vbCritical + vbOKOnly, "Alert"
     Exit Sub
  Else
     If Len(Text4.Text) < 3 Then
        MsgBox "Password Should be atleast of 4 characters", vbOKOnly + vbInformation, "Info"
        Exit Sub
     Else
        a = MsgBox("Do You want to Change Password", vbExclamation + vbOKCancel, "Confirm ?")
        If a = vbOK Then
          'change password verification
          sql = "Select * from user where userid = " & uid
          rs.Open sql, cn
            If rs!password = Text6.Text Then
               Set rs = Nothing
               sql = "update user set password = '" & Text4.Text & "' where userid = " & uid
               rs.Open sql, cn
               Set rs = Nothing
               MsgBox "New Password Updated", vbInformation + vbOKOnly, "Success"
            Else
               MsgBox "Wrong Password", vbCritical + vbOKOnly, "Error"
               Set rs = Nothing
               Exit Sub
            End If
        ElseIf a = vbNo Then
          Exit Sub
        End If
     End If
  End If
End If
Call conn
Text6.Text = Clear
Text4.Text = Clear
End Sub
Private Sub Form_Load()
'cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Port=3306;Database=examsys;User=root;password=;Option=3;"
'cn.Open
Call conn
End Sub
Private Sub conn()
sql = "Select * from user where userid = " & uid
rs.Open sql, cn
If rs.EOF = True Then
 Unload Me
 home.Show
 Exit Sub
End If
Text1.Text = rs![Name]
Text5.Text = rs!mobile
Text7.Text = rs!email
Text8.Text = rs!address
Text9.Text = rs!qans
sql = "Select * from verify where qid = " & rs!qid
Set rs = Nothing
rs.Open sql, cn
Combo2.Text = Clear
Combo2.SelText = rs!question
Set rs = Nothing
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
