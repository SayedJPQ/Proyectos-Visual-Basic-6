VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5040
   ClientLeft      =   6165
   ClientTop       =   3945
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   5040
   ScaleWidth      =   8490
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3960
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List2.Clear
For x = 0 To Val(Text1.Text) Step 3
List2.AddItem x
Next
End Sub

Private Sub Form_Load()
For x = 100 To 0 Step -2
List1.AddItem x
Next
End Sub
