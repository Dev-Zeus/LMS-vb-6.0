VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Forgot Password?"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16470
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   16470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   135
      Left            =   14040
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8760
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5880
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11280
      Picture         =   "forgot Password.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3120
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   10440
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from login"
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
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remembered Password?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   16200
      TabIndex        =   19
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(DDMMYYYY)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caps Lock is On"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11280
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10920
      TabIndex        =   14
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4560
      Width           =   6000
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter DOB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   6120
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   11160
      Picture         =   "forgot Password.frx":0719
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub Command1_Click()

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please Enter ID & DOB and then try !", vbCritical + vbOKOnly, "Try Again"
Else
Dim str As String
 str = StrComp(Adodc1.Recordset.Fields("dob").Value, Text2.Text, vbTextCompare)
 If str = True Then
 Label5.Caption = "Account not verified ,Can't reset the password"
 Label3.Caption = "Sorry .. Date of Birth Not Matched !! "
 Label8.Visible = False
 
 Label3.ForeColor = &HFF&
 Label5.ForeColor = &HFF&
 Else
 Label3.ForeColor = &H8000&
 Label5.ForeColor = &H8000&
 Label3.Caption = "Congratulations !!"
  Label8.ForeColor = &HFF&
 Label8.Caption = Adodc1.Recordset.Fields("F_name")
 Label6.Visible = True
 Label7.Visible = True
 
 Label8.ForeColor = &HFF&
 Label5.Caption = "Account is verified Now,Set your new Password"
 Text3.Visible = True
 
Text4.Visible = True
Text3.SetFocus
Command2.Visible = True
Command3.Visible = True
 Label3.Visible = True
 Label4.Visible = True
 End If
End If
End Sub

Private Sub Command2_Click()
If Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Please Enter Password", vbOKOnly + vbCritical, "Error!"
Else
If Len(Text3.Text) < 5 Then
MsgBox "Password is too Weak", vbCritical + vbOKOnly, "Error!"

Else
If Text3.Text = Text4.Text Then
Adodc1.Recordset.Fields("Password") = Text4.Text
Adodc1.Recordset.Update
MsgBox "Password Changed Successfully", vbInformation, "Password Change is Successful"
Form2.Hide
Form1.Show
Else
MsgBox "Passwords are not Identical!", vbExclamation, "Change Password:Failed"
Text3.Text = ""
Text4.Text = ""
End If
End If

End If
End Sub

Private Sub Command3_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
Label3.Caption = "Please Enter UserName To reset Password"
Label3.ForeColor = &HFF&
Else

Adodc1.RecordSource = "Select * from login where Username='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label3.Caption = "User ID Not Found ..Sorry Can't Reset the password!!!! "
Label3.ForeColor = &HFF&
Else
Label3.Caption = "User ID Found in the database"
Label3.ForeColor = &H8000&
End If
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LMS.mdb;Persist Security Info=False"
Adodc1.Visible = False
End Sub

Private Sub Image1_Click()
If Text1.Text = "" Then
Label3.Caption = "Please Enter UserName To reset Password"
Label3.ForeColor = &HFF&
Else

Adodc1.RecordSource = "Select * from login where username='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label3.Caption = "User ID Not Found ..Sorry Can't ReSet the password!!!! "
Label3.ForeColor = &HFF&
Else
Label3.Caption = "User ID Found in the database"
Label3.ForeColor = &H8000&
End If
End If
End Sub

Private Sub Label13_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then


Adodc1.RecordSource = "Select * from login where Username='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label3.Caption = "User ID Not Found ..Sorry Can't ReSet the password!!!! "
Label3.ForeColor = &HFF&
Else
Label3.Caption = "User ID Found in the database"
Label3.ForeColor = &H8000&
Text2.SetFocus
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Command1.Value = True


 
End If
End Sub

Private Sub Text3_Change()
If GetKeyState(vbKeyCapital) = 0 Then
        Label11.Visible = False
        
    Else
        Label11.Visible = True
        
    End If
    
    
    
    Dim a As Integer
a = Len(Text3.Text)


If 8 < a Then
Label10.Visible = True
Label10.ForeColor = vbGreen


Label10.Visible = True
Label10.Caption = "Password Is  strong"
Else
Label10.Visible = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)




If GetKeyState(vbKeyCapital) = 0 Then
        Label11.Visible = False
        
    Else
        Label11.Visible = True
        
    End If
    
    If KeyAscii = 13 Then
    Text4.SetFocus
    End If
    
End Sub

Private Sub Text4_Change()
If Text3.Text = "" Then
Label10.Visible = True
Label10.ForeColor = vbRed

Label10.Caption = "Please Enter Password"
Else
Label10.Visible = False
  If Text3.Text = Text4.Text Then
  Label9.Visible = True
  Label9.Caption = "Your Passwords Matched! You can Proceed Now"
  Label9.ForeColor = &H8000&
  Else
  Label9.ForeColor = vbRed
  Label9.Caption = "Passwords are not identical"
  Label9.Visible = True
  End If
End If

If GetKeyState(vbKeyCapital) = 0 Then
        Label11.Visible = False
        
    Else
        Label11.Visible = True
        
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
    End If
    




If Text3.Text = "" Then
Label10.Visible = True
Else
Label10.Visible = False
  If Text3.Text = Text4.Text Then
  Label9.Visible = True
  Label9.Caption = "Your Passwords Matched! You can Proceed Now"
  Label9.ForeColor = &H8000&
  Else
  Label9.ForeColor = vbRed
  Label9.Caption = "Passwords are not identical"
  Label9.Visible = True
  End If
End If



If GetKeyState(vbKeyCapital) = 0 Then
        Label11.Visible = False
        
    Else
        Label11.Visible = True
        
    End If
End Sub

