VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580
   LinkTopic       =   "Form3"
   ScaleHeight     =   6060
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Deshabilitar"
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Habilitar"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cajas3 False
End Sub

Private Sub Command2_Click()
Cajas3 True
End Sub

Public Sub Cajas3(Veracidad As Boolean)
Text1.Locked = Veracidad
End Sub
