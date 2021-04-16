VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   11730
   ClientTop       =   4185
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   4800
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Ejemplo 2"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Ejemplo"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text1.Text = Combo1.Text
End Sub

Private Sub Form_Click()
Form3.Show
End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "Hola"
    .AddItem "Adios"
End With

With List1
    .AddItem "Ejemplo"
    .AddItem "Sayed"
End With
End Sub

Private Sub List1_Click()
Text2.Text = List1.Text
End Sub
