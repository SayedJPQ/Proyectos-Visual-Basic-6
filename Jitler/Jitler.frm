VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Jitler Click"
   ClientHeight    =   5490
   ClientLeft      =   3705
   ClientTop       =   2805
   ClientWidth     =   11985
   FillColor       =   &H00800000&
   ForeColor       =   &H0000FF00&
   Icon            =   "Jitler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   11985
   Begin VB.CommandButton Command6 
      Caption         =   "Cerrar Jitler"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pausar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reiniciar"
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver resultado"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8520
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Iniciar"
      Height          =   735
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jitler click"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Clicks por segundo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clicks As Double
Dim s As Double
Dim op As String

Private Sub Command1_Click()
clicks = clicks + 1
Label1 = clicks
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Text1.Text = (clicks)
op = "/"
If op = "/" Then
Text1.Text = clicks / s
End If
End Sub

Private Sub Command4_Click()
clicks = 0
Text1.Text = ""
Text1.SetFocus
s = 0
Label1 = 0
Label2 = 0
Command2.Enabled = True
End Sub

Private Sub Command5_Click()
Timer1.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Timer1_Timer()
s = s + 1
Label2 = s
If s = 10 Then
Timer1.Enabled = False
End If
If s = 10 Then
Command1.Enabled = False
If s = 10 Then
Command2.Enabled = False
End If
End If
End Sub
