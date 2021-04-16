VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6150
   ClientLeft      =   4005
   ClientTop       =   2775
   ClientWidth     =   13455
   LinkTopic       =   "Form3"
   ScaleHeight     =   6150
   ScaleWidth      =   13455
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ingrese su usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim speaks, speech
Set speech = CreateObject("sapi.spvoice")
Dim Usuario
Dim Respuesta
Usuario = "sayed"
If Text1.Text = "sayed" Then
Respuesta = "Ahora debes poner tu contraseña"
speaks = Respuesta
speech.Speak speaks
Form3.Hide
Form2.Show
Else
Respuesta = "No estas autorizado"
speaks = Respuesta
speech.Speak speaks
MsgBox ("Vete de aqui")
End
End If
End Sub

