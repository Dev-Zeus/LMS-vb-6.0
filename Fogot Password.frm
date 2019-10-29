VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   " Forgot Pasword ?"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   10560
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LMS\LMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LMS\LMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from login"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Text            =   " "
      Top             =   6840
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Text            =   " "
      Top             =   6000
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   495
      Left            =   11400
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Text            =   " "
      Top             =   4560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Search"
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Text            =   " "
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   " Remembered Password ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   17760
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Comfirm Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   " Accoount is verified Now,Set your new password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   " (DDMMYYYY)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   " Enter DOB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   " Enter Username  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   " Forgot Password ?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()

End Sub

Private Sub Label8_Click()
Form1.Show
Form2.Hide

End Sub
