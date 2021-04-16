VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6735
   ClientLeft      =   4005
   ClientTop       =   2250
   ClientWidth     =   11145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6735
   ScaleWidth      =   11145
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   4800
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   4200
      Width           =   4695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   4440
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   7320
      Top             =   1080
      Width           =   3495
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
      Left            =   120
      TabIndex        =   8
      Top             =   4800
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
      Left            =   120
      TabIndex        =   6
      Top             =   4200
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
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
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
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   1095
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
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Reproductor de imagenes"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
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
Image1.Picture = LoadPicture(File1.Path & "/" & File1.FileName)
End Sub

