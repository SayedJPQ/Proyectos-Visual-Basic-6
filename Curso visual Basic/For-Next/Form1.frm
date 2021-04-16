VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   4860
   ClientTop       =   2850
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   10890
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   2280
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregrar"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Inserte numero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Numeros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List2.Clear
For x = 0 To Val(Text1.Text)
List2.AddItem x
Next
End Sub

Private Sub Command2_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub Form_Load()
For x = 0 To 100
List1.AddItem x
Next
End Sub

