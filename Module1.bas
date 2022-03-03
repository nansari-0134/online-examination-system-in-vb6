Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Public uid As Long
Public ut As Integer
Public qno As Integer
Public tid As Integer
Public tm As Double
Public q() As Integer
Public cnt As Integer
Public time As Long
Sub main()
cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Port=3306;Database=examsys;User=root;password=;Option=3;"
cn.Open
login.Show
tm = 0
End Sub
  
