VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3255
   ClientLeft      =   4905
   ClientTop       =   1485
   ClientWidth     =   3075
   LinkTopic       =   "Form3"
   ScaleHeight     =   3255
   ScaleWidth      =   3075
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Combo1.RemoveItem (Combo1.ListIndex)
End Sub

Private Sub Command2_Click()
List1.RemoveItem (List1.ListIndex)
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
