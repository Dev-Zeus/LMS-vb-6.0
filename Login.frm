VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":0000
   ScaleHeight     =   9885
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   8640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Text            =   " "
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Text            =   " "
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      BorderStyle     =   4  'Dash-Dot
      Height          =   6495
      Left            =   3000
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   8040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Forgot Password ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter Username", vbOKOnly + vbCritical, "Error!"
Else
If Text2.Text = "" Then
MsgBox "Please Enter Password", vbOKOnly + vbCritical, "Error!"
Text2.SetFocus
Else
Adodc1.RecordSource = "select * from login where username='" + Text1.Text + "' and password='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Login failed,Try Again..!!!", vbCritical, "Please Enter correct Username and Password"
Text2.Text = ""
Text2.SetFocus
Else
Form8.Show
Form1.Hide
End If
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"

End Sub

Private Sub Text1_LostFocus()
Dim adovv As New ADODB.Connection
Dim Rado As New ADODB.Recordset
Dim constr As String
Dim a As String
a = Form1.Text1.Text
constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adovv.ConnectionString = constr
adovv.Open
Rado.Source = "SELECT * FROM login where [Username]='" & a & "'"
Rado.CursorType = adOpenForwardOnly
Rado.ActiveConnection = adovv
Rado.Open
Do While Not Rado.EOF
     
     Label5.Caption = Rado.Fields("F_name").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
End Sub
