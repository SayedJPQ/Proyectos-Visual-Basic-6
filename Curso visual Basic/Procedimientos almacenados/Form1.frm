VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   4650
   ClientTop       =   1980
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11460
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Hola6"
      Top             =   4440
      Width           =   6615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Text            =   "Hola5"
      Top             =   3720
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "Hola4"
      Top             =   3000
      Width           =   6615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Text            =   "Hola3"
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Hola2"
      Top             =   1560
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Text            =   "Hola1"
      Top             =   840
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Cajas()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Form_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub Form_Load()
Cajas
End Sub

