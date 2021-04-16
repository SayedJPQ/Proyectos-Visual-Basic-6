VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   6345
   ClientTop       =   1995
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7920
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   4680
      Pattern         =   "*.jpg;*.jpeg;*.gif;*.bmp;*"
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   6000
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   5160
      Width           =   4695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Directorio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Archivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ruta del archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Informacion del archivo"
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
      TabIndex        =   3
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Explorador de archivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Path
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
Text2.Text = File1.Path & "/" & File1.FileName
End Sub

Private Sub Form_Click()
Form2.Show
End Sub

