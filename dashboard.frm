VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15840
   LinkTopic       =   "Form9"
   ScaleHeight     =   8955
   ScaleWidth      =   15840
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome!"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = Form1.Label5.Caption

End Sub
