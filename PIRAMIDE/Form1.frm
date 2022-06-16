VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   12390
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   ScaleHeight     =   12390
   ScaleWidth      =   16020
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   5280
      TabIndex        =   7
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   6
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   9480
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "EL VOLUMEN ES IGUAL A:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "INTRODUSCA LA ALTURA :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "INTRODUSCA EL AREA DE LA BASE:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "PIRAMIDE:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Image1.Picture = LoadPicture("E:\PIRAMIDE\PIRA.jpg")

Dim AREA As Double
Dim ALTURA As Double
Dim VOLUMEN As Double

AREA = Val(Text1.Text)
ALTURA = Val(Text2.Text)

VOLUMEN = (AREA * ALTURA / 3)
Label5 = VOLUMEN
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Label5 = ""

Image1.Picture = Nothing


End Sub

