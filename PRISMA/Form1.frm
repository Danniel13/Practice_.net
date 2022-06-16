VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   12645
   Begin VB.TextBox Text5 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   12000
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000009&
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "salir"
      Height          =   855
      Left            =   7440
      TabIndex        =   9
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "limpiar"
      Height          =   855
      Left            =   4200
      TabIndex        =   8
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calcular"
      Height          =   855
      Left            =   1200
      TabIndex        =   7
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8160
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "volumen del prisma "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   15
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "altura del prisma "
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
      Height          =   615
      Left            =   9840
      TabIndex        =   13
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "area de  base"
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
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "numero de lados"
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
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11880
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "area de lado "
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9960
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "altura de lado"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "base de lado  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "VOLUMEN DE UN PRISMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim BASEDELADO As Double
Dim ALTURADELADO As Double
Dim AREADEUNLADO As Double
Dim NUMERODELADO As Double
Dim VOLUMEN As Double
 
BASEDELADO = Val(Text1.Text)
ALTURADELADO = Val(Text2.Text)
AREADELADO = (BASEDELADO * ALTURADELADO) / 2
Label5 = AREADELADO

AREADELADO = Val(Label5)
NUMERODELADOS = Val(Text3.Text)
AREADEBASE = (AREADELADO * NUMERODELADOS)
Label10 = AREADEBASE

AREADEBASE = Val(Label10)
ALTURADELPRISMA = Val(Text5.Text)
VOLUMENDELPRISMA = (AREADEBASE * ALTURADELPRISMA)
Label11 = VOLUMENDELPRISMA

End Sub

Private Sub Command2_Click()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Label5 = ""
Label10 = ""
Label11 = ""
End Sub

Private Sub Command3_Click()
End

End Sub

