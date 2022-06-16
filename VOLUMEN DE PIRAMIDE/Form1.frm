VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   14
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5520
      TabIndex        =   13
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5520
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "SALIR"
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "BORRAR"
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "CALCULAR"
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   15
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "VOLUMEN PIRAMIDE"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ALTURA PIRAMIDE"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ALTURA LADO"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "BASE DE UN LADO"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NUMERO DELADOS"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "AREA DE UN LADO"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VOLUMEN TOTAL  DE UNA PIRAMIDE"
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
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ALTURALADO As Double
Dim NUMERODELADOS As Double
Dim AREALADO As Double
Dim BASELADO As Double
Dim ALTURAPIRAMIDE As Double
Dim PERIMETRO As Double
ALTURALADO = Val(Text1.Text)
NUMERODELADOS = Val(Text3.Text)
AREALADO = Val(Text3.Text)
BASELADO = Val(Text4.Text)
ALTURAPIRAMIDE = Val(Text5.Text)
VOLUMENPRISMA = (ALTURALADO*PERIMETRO/)
Label8 = VOLUMENPRISMA

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Label8 = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub
