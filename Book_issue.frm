VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Book Issue                                                        "
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6960
      TabIndex        =   19
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   6960
      TabIndex        =   18
      Top             =   3120
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   14760
      TabIndex        =   16
      Text            =   "Select"
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   645
      Left            =   11640
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   15840
      TabIndex        =   14
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Issue"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   8640
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   14760
      TabIndex        =   10
      Text            =   "Select"
      Top             =   6120
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
      Left            =   7080
      TabIndex        =   9
      Text            =   " "
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      Left            =   7080
      TabIndex        =   8
      Text            =   " "
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox Text3 
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
      Left            =   7080
      TabIndex        =   7
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Student Name"
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
      Left            =   12000
      TabIndex        =   17
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   " OR"
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
      Left            =   10680
      TabIndex        =   11
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Date "
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
      Left            =   4080
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   " Book Name"
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
      Left            =   12120
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Book ID"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Return Date"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Transaction ID"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Student Id"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Book Issue"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Dim adovv As New ADODB.Connection
Dim Rado As New ADODB.Recordset
Dim constr As String
Dim a As String
a = Form1.Text1.Text
constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adovv.ConnectionString = constr
adovv.Open
Rado.Source = "SELECT * FROM student where [St_name]='" & a & "'"
Rado.CursorType = adOpenForwardOnly
Rado.ActiveConnection = adovv
Rado.Open
Do While Not Rado.EOF
     
     Label5.Caption = Rado.Fields("BL").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
End Sub

Private Sub Combo2_Click()
Dim adovv As New ADODB.Connection
Dim Rado As New ADODB.Recordset
Dim constr As String
Dim a As String
a = Combo2.Text
constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adovv.ConnectionString = constr
adovv.Open
Rado.Source = "SELECT * FROM student where [St_name]='" & a & "'"
Rado.CursorType = adOpenForwardOnly
Rado.ActiveConnection = adovv
Rado.Open
Do While Not Rado.EOF
     
     Text5.Text = Rado.Fields("BL").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
End Sub

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

 strSQL = "INSERT INTO tran([trn_id],[stud_name],[book_name],[b_id],[fdate],[rdate]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text4.Text & "','" & Text6.Text & "' )"
 

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
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Combo1.Text = "Select"
End Sub

Private Sub Command3_Click()
Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strConn1 As String
strConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strConn1
adoConn1.Open
adoRS1.Source = "SELECT * FROM tran"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Text1.Text = adoRS1.Fields("trn_id").Value + 1
     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Set adoConn1 = Nothing
End Sub

Private Sub Form_Load()
Dim Edate, bdate As Date
'Edate = CInt(Text6.Text)
'Text3.Text = bdate
'bdate = CInt(Edate) + CInt(7)


Command3.Value = True

Combo1.Text = "Select "
Text6.Text = Format(Now, "dd/mm/yyyy")
'for combobox1
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
     Combo1.AddItem (adoRS1.Fields("book_name").Value)

     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Set adoConn1 = Nothing
'for combobox2
Dim adoConn2 As New ADODB.Connection
Dim adoRS2 As New ADODB.Recordset
Dim strConn2 As String
strConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
adoConn2.ConnectionString = strConn2
adoConn2.Open
adoRS2.Source = "SELECT * FROM student"
adoRS2.CursorType = adOpenForwardOnly
adoRS2.ActiveConnection = adoConn2
adoRS2.Open
Do While Not adoRS2.EOF
     Combo2.AddItem (adoRS2.Fields("St_name").Value)

     adoRS2.MoveNext
Loop
adoRS2.Close
Set adoRS2 = Nothing
adoConn2.Close
Set adoConn2 = Nothing

End Sub

Private Sub Text5_Change()
If Text5.Text = "FALSE" Then
Command1.Enabled = True

Else
MsgBox "Student is blocked Please Ask to return books", vbCritical + vbOKOnly, "Error!"
Command1.Enabled = False

End If
'Combo2.Clear




End Sub

