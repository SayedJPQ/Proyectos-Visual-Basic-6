VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5010
   ClientLeft      =   4530
   ClientTop       =   3285
   ClientWidth     =   12030
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   12030
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar sesión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000005&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   0
      Top             =   2880
      Width           =   7935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ingrese su contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   8775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Inicio sesion
Dim speaks, speech
Set speech = CreateObject("sapi.spvoice")
Dim Contraseña
Dim NOMBRE
Dim Respuesta As String
NOMBRE = Text1.Text
Contraseña = "sayed123"
If NOMBRE = Contraseña Then
Form2.Hide
Form1.Show
Respuesta = "Bienvenido señor, le mostrare una lista de comandos que puede ejecutar y una guía"
speaks = Respuesta
speech.Speak speaks
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "UserAccountControlSettings.exe"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "C:\Users\PC0\Desktop\Kara\Tutorial\Tutorial.exe"
Else
Respuesta = "No estas autorizado"
speaks = Respuesta
speech.Speak speaks
MsgBox "Vete de aqui"
End
End If
End Sub


