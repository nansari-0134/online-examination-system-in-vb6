VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form addque 
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   9135
      Left            =   0
      TabIndex        =   41
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
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
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
      Left            =   4440
      TabIndex        =   29
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Add Questions"
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
      Height          =   7215
      Left            =   4440
      TabIndex        =   9
      Top             =   2520
      Width           =   15855
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
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
         Left            =   14520
         TabIndex        =   42
         Top             =   3530
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Upload"
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
         TabIndex        =   30
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1080
         TabIndex        =   25
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   6240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "addque.frx":0000
         Left            =   1080
         List            =   "addque.frx":0010
         TabIndex        =   22
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   12480
         TabIndex        =   19
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   8400
         TabIndex        =   17
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4560
         TabIndex        =   15
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   4080
         Width           =   15495
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   5741
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qno."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Question"
            Object.Width           =   25003
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "option1"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "option2"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "option3"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "option4"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "answer"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Mark"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks :  "
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   6890
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   0
         Top             =   6840
         Width           =   16335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mark :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "D."
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12120
         TabIndex        =   20
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "C."
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   18
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "B."
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   16
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "A."
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Question :"
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Add New Test"
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
      Height          =   1935
      Left            =   4440
      TabIndex        =   4
      Top             =   530
      Width           =   15855
      Begin MSComCtl2.DTPicker text10 
         Height          =   375
         Left            =   7800
         TabIndex        =   44
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   119734273
         CurrentDate     =   43768
      End
      Begin VB.CommandButton Command7 
         Caption         =   "New"
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
         Left            =   4320
         TabIndex        =   43
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   39
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1320
         Width           =   14295
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13800
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   28
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   27
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   3015
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
         Left            =   11640
         TabIndex        =   40
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label17 
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
         TabIndex        =   38
         Top             =   1320
         Width           =   855
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
         Left            =   14520
         TabIndex        =   36
         Top             =   360
         Width           =   495
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
         Left            =   13320
         TabIndex        =   34
         Top             =   360
         Width           =   375
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
         Left            =   11640
         TabIndex        =   33
         Top             =   360
         Width           =   855
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
         Left            =   6000
         TabIndex        =   8
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
         TabIndex        =   7
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
         Left            =   6000
         TabIndex        =   6
         Top             =   360
         Width           =   1095
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
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
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
      Caption         =   "Add / Remove Tests"
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
Attribute VB_Name = "addque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim n As Integer
Dim tm As Long
Private Sub Command1_Click()
Unload Me
home.Show
End Sub
Private Sub Command2_Click()
Set rs = Nothing
For i = ListView1.ListItems.Count To 1 Step -1
    ''''''
    If ListView1.ListItems(i).Checked = True Then
    'delete from test
     sql = "delete from test where testid =" & ListView1.ListItems(i)
     rs.Open sql, cn
     Set rs = Nothing
    'delete from question
    sql = "delete from question where testid = " & ListView1.ListItems(i)
    rs.Open sql, cn
    Set rs = Nothing
   End If
Next i
Unload Me
addque.Show
End Sub
Private Sub Command4_Click()
Command6.Enabled = False
Unload Me
addque.Show
End Sub
Private Sub Command5_Click()
If validate = True Then
        If CInt(Text11.Text) - 2 >= ListView2.ListItems.Count Then
           'add in listview
          Set it = ListView2.ListItems.Add(, , ListView2.ListItems.Count + 1)
            it.SubItems(1) = Text1.Text
            it.SubItems(2) = Text2.Text
            it.SubItems(3) = Text3.Text
            it.SubItems(4) = Text4.Text
            it.SubItems(5) = Text5.Text
            it.SubItems(6) = Combo1.Text
            it.SubItems(7) = Text6.Text
            Text1.Text = Clear
            Text2.Text = Clear
            Text3.Text = Clear
            Text4.Text = Clear
            Text5.Text = Clear
            Combo1.Text = Clear
            Text7.Text = ListView2.ListItems.Count + 1
            'totalmark
            tm = 0
            For i = ListView2.ListItems.Count To 1 Step -1
                tm = tm + CLng(ListView2.ListItems(i).SubItems(7))
            Next
            Label13.Caption = "Total Marks : " & tm
            
        ElseIf CInt(Text11.Text) - 1 <= ListView2.ListItems.Count Then
            Set it = ListView2.ListItems.Add(, , ListView2.ListItems.Count + 1)
            it.SubItems(1) = Text1.Text
            it.SubItems(2) = Text2.Text
            it.SubItems(3) = Text3.Text
            it.SubItems(4) = Text4.Text
            it.SubItems(5) = Text5.Text
            it.SubItems(6) = Combo1.Text
            it.SubItems(7) = Text6.Text
            MsgBox "Uploading test ..." & vbNewLine & "Total Number Of questions : " & n & ".", vbInformation + vbOKOnly, "Notification"
            'upload all ..
            Set rs = Nothing
            If Text12.Text = "" Then
               Text12.Text = "0"
            End If
            If Text13.Text = "" Then
               Text13.Text = "0"
            End If
            sql = "insert into test(testid,tdate,ttime,course,noq,topic,detail) values(" & Text8.Text _
                  & ",'" & Format(Date, "yyyy-mm-dd") & "'," & CLng(Text12.Text) * 60 + CLng(Text13.Text) _
                  & ",'" & Text9.Text & "'," & CLng(Text11.Text) & ",'" & Text15.Text _
                  & "','" & Text14.Text & "')"
            rs.Open sql, cn
            Set rs = Nothing
            For i = ListView2.ListItems.Count To 1 Step -1
                sql = "insert into question(qno,que,op1,op2,op3,op4,ans,mark,testid) values(" & ListView2.ListItems(i) _
                & ",'" & ListView2.ListItems(i).SubItems(1) & "'" _
                & ",'" & ListView2.ListItems(i).SubItems(2) & "'" _
                & ",'" & ListView2.ListItems(i).SubItems(3) & "'" _
                & ",'" & ListView2.ListItems(i).SubItems(4) & "'" _
                & ",'" & ListView2.ListItems(i).SubItems(5) & "'" _
                & ",'" & ListView2.ListItems(i).SubItems(6) & "'" _
                & "," & ListView2.ListItems(i).SubItems(7) & "," & Text8.Text & ")"
                rs.Open sql, cn
                Set rs = Nothing
            Next
            MsgBox "Test Uploaded Successfully", vbInformation + vbOKOnly, "Notification"
            Command5.Enabled = False
            Command7.Enabled = True
            Command2.Enabled = True
            Command6.Enabled = False
            Unload Me
            addque.Show
        End If
End If
End Sub
Private Sub Command6_Click()
For i = ListView2.ListItems.Count To 1 Step -1
    If ListView2.ListItems(i).Checked = True Then
        ListView2.ListItems.Remove (i)
    End If
Next i
For i = 1 To ListView2.ListItems.Count
     ListView2.ListItems(i).Text = i
Next i
End Sub
Private Sub Command7_Click()
ListView2.ListItems.Clear
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = Clear
Text7.Text = Clear
Combo1.Text = Clear
Text8.Text = Clear
Text9.Text = Clear
text10.Value = Clear
Text11.Text = Clear
Text12.Text = Clear
Text13.Text = Clear
Text14.Text = Clear
Text15.Text = Clear
Set rs = Nothing
sql = "select max(testid) from test"
rs.Open sql, cn
Text8.Text = rs.Fields(0)
Text8.Text = CLng(Text8.Text) + 1
Text7.Text = "1"
Set rs = Nothing
text10.Value = Date
Command7.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
Command2.Enabled = False
Label13.Caption = "Total Marks : 0"
End Sub

Private Sub Form_Load()
Call data
Command5.Enabled = False
Command6.Enabled = False
End Sub
Private Sub data()
Dim c As String
'listview1: tests(remove)
sql = "select * from user where userid = " & uid
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
If rs!course = "Admin" Or rs!course <> "" Then
    Set rs = Nothing
    sql = "select * from test"
    rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Else
    c = rs!course
    Set rs = Nothing
    sql = "select * from test where course = '" & c & "'"
    rs.Open sql, cn, adOpenDynamic, adLockOptimistic
End If
While Not rs.EOF
    Set it = ListView1.ListItems.Add(, , rs!testid)
    it.SubItems(1) = rs!topic
    it.SubItems(2) = rs!tdate
    rs.MoveNext
Wend

Command5.Enabled = False

Set rs = Nothing
'sql = "select max(testid) from test"
'rs.Open sql, cn
'Text8.Text = rs.Fields(0)
'Text8.Text = CLng(Text8.Text) + 1
'Set rs = Nothing
'text10.value = Date
End Sub
Private Sub ListView1_Click()
ListView2.ListItems.Clear
sql = "select * from test where testid = " & ListView1.SelectedItem
Set rs = Nothing
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Text8.Text = rs!testid
Text9.Text = rs!course
text10.Value = rs!tdate
Text11.Text = rs!noq
Text12.Text = rs!ttime / 60
Text13.Text = rs!ttime Mod 60
Text14.Text = IIf(IsNull(rs!detail), " ", rs!detail)
Text15.Text = rs!topic
Set rs = Nothing
sql = "select * from question where testid = " & ListView1.SelectedItem
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   Set it = ListView2.ListItems.Add(, , rs!qno)
   it.SubItems(1) = rs!que
   it.SubItems(2) = rs!op1
   it.SubItems(3) = rs!op2
   it.SubItems(4) = rs!op3
   it.SubItems(5) = rs!op4
   it.SubItems(6) = rs!ans
   it.SubItems(7) = rs!mark
   rs.MoveNext
Wend
'totalmark
    tm = 0
    For i = ListView2.ListItems.Count To 1 Step -1
        tm = tm + CLng(ListView2.ListItems(i).SubItems(7))
    Next
    Label13.Caption = "Total Marks : " & tm
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text11_LostFocus()
On Error GoTo l:
If CLng(Text11.Text) > ListView2.ListItems.Count Then
   n = CLng(Text11.Text)
ElseIf CLng(Text11.Text) < ListView2.ListItems.Count Then
   MsgBox "Number of Question Uploaded is Greater then Number of Question ", vbOKOnly + vbCritical, "Change No. Of Question"
ElseIf Text11.Text = ListView2.ListItems.Count Then
       MsgBox "Uploading test ..." & vbNewLine & "Total Number Of questions : " & n & ".", vbInformation + vbOKOnly, "Notification"
    'upload all ..
End If
Exit Sub
l:
End Sub


Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii < 58) Or (KeyAscii = 8) Or (KeyAscii = 46) Then
   KeyAscii = KeyAscii
Else
   KeyAscii = 0
End If
End Sub
Private Function validate() As Boolean
If Text1.Text = "" Then
   MsgBox "Please enter question", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text2.Text = "" Then
   MsgBox "Please enter Option A", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text3.Text = "" Then
   MsgBox "Please enter option B", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text4.Text = "" Then
   MsgBox "Please enter Option C", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text5.Text = "" Then
   MsgBox "Please enter option D", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Combo1.Text = "" Then
   MsgBox "Please select right option", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
If Text6.Text = "" Then
   MsgBox "Please enter mark", vbCritical + vbOKOnly, "Error"
   validate = False
   Exit Function
End If
validate = True
End Function
