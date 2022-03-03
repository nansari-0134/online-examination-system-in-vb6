VERSION 5.00
Begin VB.Form que 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Ubuntu"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9360
      Top             =   10680
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   10680
      TabIndex        =   22
      Top             =   5520
      Width           =   9255
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   2775
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mark Question"
      Height          =   615
      Left            =   17520
      TabIndex        =   21
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   615
      Left            =   14110
      TabIndex        =   20
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save And Next"
      Height          =   615
      Left            =   10680
      TabIndex        =   19
      Top             =   8760
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   8280
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   8280
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   20295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18240
      TabIndex        =   24
      Top             =   9840
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17280
      TabIndex        =   18
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Quetion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   9840
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   9600
      Width           =   21255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "D."
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
      Left            =   5520
      TabIndex        =   16
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "B."
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
      Left            =   240
      TabIndex        =   15
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "C."
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
      Left            =   5520
      TabIndex        =   14
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "A."
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
      Left            =   240
      TabIndex        =   13
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5880
      TabIndex        =   12
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5400
      TabIndex        =   11
      Top             =   7440
      Width           =   5055
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5880
      TabIndex        =   9
      Top             =   5400
      Width           =   4455
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5400
      TabIndex        =   8
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   6
      Top             =   7560
      Width           =   4455
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   5055
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Ubuntu"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   3
      Top             =   5400
      Width           =   4455
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   5055
   End
End
Attribute VB_Name = "que"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rsa As New ADODB.Recordset
Private Sub Command1_Click()
If Option1.Value = True Then
   If rsa!ans = 1 Then
         tm = tm + rsa!mark
   End If
ElseIf Option2.Value = True Then
   If rsa!ans = 2 Then
         tm = tm + rsa!mark
   End If
ElseIf Option3.Value = True Then
   If rsa!ans = 3 Then
         tm = tm + rsa!mark
   End If
ElseIf Option4.Value = True Then
   If rsa!ans = 4 Then
         tm = tm + rsa!mark
   End If
Else
    MsgBox "Choose any option first", vbInformation + vbOKOnly, "Save & next "
    Exit Sub
End If
Unload Me
tests.Command5.Value = True
Set rsa = Nothing
End Sub

Private Sub Command2_Click()
cnt = cnt + 1
ReDim Preserve q(cnt)
q(cnt - 1) = rsa!qno
Unload Me
tests.Command5.Value = True
Set rsa = Nothing
End Sub

Private Sub Command3_Click()
Dim s As New ADODB.Recordset
sql = "insert into mque(testid,uid,qno) values(" & tid & "," & uid & "," & rsa!qno & ")"
s.Open sql, cn
Set s = Nothing
Unload Me
tests.Command5.Value = True
Set rsa = Nothing
End Sub

Private Sub Form_Load()
Dim r As New ADODB.Recordset
   Label15.Caption = vbNewLine & "Note :-" & vbNewLine & vbNewLine & "Save and next : It will save your answer" & vbNewLine & vbNewLine & "Next : it will move to the next question" _
                        & " And will show this question again after the end of all questions" & vbNewLine & vbNewLine & "Mark Question : it will completly skip this question and will Give you the Solution after the completition of test"
sql = "Select * from question where testid = " & tid & " and qno = " & qno
Set rsa = Nothing
rsa.Open sql, cn, adOpenDynamic, adLockOptimistic
Text1.Text = rsa!que
Label2.Caption = rsa!op1
Label4.Caption = rsa!op2
Label6.Caption = rsa!op3
Label8.Caption = rsa!op4
sql = "select count(*) from question where testid = " & tid
r.Open sql, cn
Label13.Caption = "Question : " & rsa!qno & " / " & r.Fields(0)
Set r = Nothing
End Sub
Private Sub Label2_Click()
If Option1.Value = True Then
    Option1.Value = False
Else
    Option1.Value = True
End If

End Sub
Private Sub Label4_Click()
If Option2.Value = True Then
    Option2.Value = False
Else
    Option2.Value = True
End If
End Sub
Private Sub Label6_Click()
If Option3.Value = True Then
    Option3.Value = False
Else
    Option3.Value = True
End If
End Sub
Private Sub Label8_Click()
If Option4.Value = True Then
    Option4.Value = False
Else
    Option4.Value = True
End If
End Sub
Private Sub Timer1_Timer()
Dim h, m As Integer
If time <> 0 Then
   time = time - 1
   h = time \ 3600
   m = (time - h * 3600) \ 60
   Label16.Caption = h & " : " & m & " : " & time Mod 60
Else
   Unload Me
   tests.Command4.Value = True
End If
End Sub
