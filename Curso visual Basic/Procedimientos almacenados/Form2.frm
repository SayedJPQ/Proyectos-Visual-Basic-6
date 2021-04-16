VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11460
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Form_Load()
Cajas2
End Sub

Public Sub Cajas2()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
End Sub

