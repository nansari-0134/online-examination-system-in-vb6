VERSION 5.00
Begin VB.MDIForm mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Online Exam System"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
home.Show
End Sub
