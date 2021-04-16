VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   2175
   ClientTop       =   7245
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   5835
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Combo1.AddItem Text1.Text
Text1.Text = ""
End Sub

Private Sub Command2_Click()
List1.AddItem Text2.Text
Text2.Text = ""
End Sub

Private Sub Form_Click()
Form2.Show
End Sub

Private Sub List1_Click()

End Sub
