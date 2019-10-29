VERSION 5.00
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Return Window"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9000
      TabIndex        =   13
      Text            =   " "
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9000
      TabIndex        =   11
      Text            =   " "
      Top             =   7200
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   17640
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   13080
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Date Taken"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Book id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Return Window"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim adovv As New ADODB.Connection
Dim Rado As New ADODB.Recordset
Dim constr As String
Dim a As String
a = Val(Text1.Text)
constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adovv.ConnectionString = constr
adovv.Open
Rado.Source = "SELECT * FROM book where [b_id]='" & a & "'"
Rado.CursorType = adOpenForwardOnly
Rado.ActiveConnection = adovv
Rado.Open
Do While Not Rado.EOF
     
     Text2.Text = Rado.Fields("book_name").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
   

End Sub

Private Sub Form_Load()
Text6.Text = Format(Now, "dd/mm/yyyy")
End Sub

