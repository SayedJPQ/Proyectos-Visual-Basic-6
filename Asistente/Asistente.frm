VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kara"
   ClientHeight    =   8145
   ClientLeft      =   2445
   ClientTop       =   1965
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Asistente.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Asistente.frx":1084A
   ScaleHeight     =   8145
   ScaleWidth      =   16140
   Begin VB.CommandButton Command5 
      Caption         =   "Finalizar programa"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14040
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Traducir"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Traducido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10920
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Traducir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Leer en voz alta"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Leer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox Respuesta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Pregunta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Leer texto a voz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "BOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Traductor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CodigoElse(texto As String, Traduce As Boolean)
Dim speaks, speech
Set speech = CreateObject("sapi.spvoice")
If Traduce = False Then
Select Case texto

'Interaccion con el bot

Case "Hola"
Respuesta.Text = texto
Case "Como estas?"
Respuesta.Text = "Bien y usted señor"
Case "Me siento triste"
Respuesta.Text = "Espero que te sientas mejor. Pondre algo de musica"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "https://www.youtube.com/playlist?list=LLsnM9cuc8SNfT9N3lrscWRw"
Case "Me siento feliz"
Respuesta.Text = "Me alegro por usted señor"
Case "Quiero programar"
Respuesta.Text = "En que lenguaje desea programar señor?"

'Abrir programas
Case "Python"
Respuesta.Text = "Python abierto señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "python.exe"
Case "CMD"
Respuesta.Text = "CMD abierto señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "cmd.exe"
Case "Inventario"
Respuesta.Text = "Programa de inventario abierto señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "C:\Users\PC0\Desktop\VisualProjects\Inventario\Inventario.exe"

'Abrir webs
Case "Youtube"
Respuesta.Text = "youtube abierto señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "www.youtube.com"
Case "Quiero escuchar musica"
Respuesta.Text = "¿Que tipo de musica te gustaria escuchar?"
Case "Quiero escuchar rock"
Respuesta.Text = "Buscando musica rock señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "https://www.youtube.com/results?search_query=musica+rock"
Case "Facebook"
Respuesta.Text = "Facebook abierto señor"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "www.facebook.com"

'Tutorial
Case "Que puedo traducir"
Respuesta.Text = "Se presentara a continuacion la lista de palabras que puede traducir, por favor pongalas tal y como estan para evitar errores"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "C:\Users\PC0\Desktop\Kara\Translate\Translate.exe"

'Cerrar programa
Case "Cerrar programa"
MsgBox "Hasta luego"
End

Case Else
Respuesta.Text = "NO LE ENTENDI SEÑOR"
End Select
speaks = Respuesta.Text
speech.Speak speaks

'Traducir
Else
Select Case texto
'Saludos y despedidas
Case "Saludo"
Traducido.Text = "Greeting"
Case "Despedida"
Traducido.Text = "Farewell"
Case "Hola"
Traducido.Text = "Hello"
Case "Como estas?"
Traducido.Text = "How are you"
Case "Adios"
Traducido.Text = "Goodbye"
Case "Te veo despues"
Traducido.Text = "See you later"
Case "Te veo pronto"
Traducido.Text = "See you soon"
'Objetos
Case "Libro"
Traducido.Text = "Book"
Case "Lapicero"
Traducido.Text = "Pen"
Case "Lapiz"
Traducido.Text = "Pencil"
Case Else
Traducido.Text = "NO LE ENTENDI SEÑOR"
End Select
speaks = Traducido.Text
speech.Speak speaks
End If
End Sub

Private Sub Command1_Click()
CodigoElse Pregunta.Text, False
End Sub

Private Sub Command3_Click()
Dim speaks, speech
speaks = (Leer.Text)
Set speech = CreateObject("sapi.spvoice")
speech.Speak speaks
End Sub
Private Sub Command4_Click()
CodigoElse Traducir.Text, True
End Sub

Private Sub Command5_Click()
MsgBox "Hasta luego"
End
End Sub

