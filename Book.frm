VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Book"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form4"
   ScaleHeight     =   7350
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6720
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Text            =   " "
      Top             =   5760
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Text            =   " "
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   9
      Text            =   " "
      Top             =   3600
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Text            =   " "
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Text            =   " "
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label7 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Publication"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Add Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Are you sure want to add this Record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
Else


 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO book([b_id],[book_name],[author],[pub],[yearp],[amount]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Combo1.Text & "','" & Text4.Text & "','" & Text6.Text & "')"
 

 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute
 Text2.Text = " "
  Text3.Text = " "
  Text4.Text = " "
  Text6.Text = " "
MsgBox "Book Details Added Successfully", vbOKOnly + vbInformation, "Done!"
 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
ICanDealWithThis:
 MsgBox "Something went wrong!,Please try filling all details properly", vbCritical + vbOKOnly, "Error!"
End If
End Sub

Private Sub Command2_Click()
Text1.Text = Val(Text1.Text) + 1
  Text2.Text = " "
  Text3.Text = " "
  Text4.Text = " "
  Text6.Text = " "
End Sub

Private Sub Form_Load()

Combo1.Text = "Select"
Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strConn1 As String
strConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strConn1
adoConn1.Open
adoRS1.Source = "SELECT * FROM book"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Text1.Text = adoRS1.Fields("b_id").Value + 1
     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Combo1.Clear


Set adoConn1 = Nothing

Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM publication"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo1.AddItem (adoRS.Fields("p_name").Value)

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing


End Sub

