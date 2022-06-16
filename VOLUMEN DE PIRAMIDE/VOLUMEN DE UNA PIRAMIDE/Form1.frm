VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SALIR"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BORRAR"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CALCULAR"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   7560
      Picture         =   "Form1.frx":0000
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   7440
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "VOLUMEN TOTAL"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ALTURA"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "AREA DE LA BASE"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "VOLUMEN DE UNA PIRAMIDE"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AREABASE As Double
Dim ALTURA As Double

AREABASE = Val(Text1.Text)
ALTURA = Val(Text2.Text)
VOLUMENTOTAL = (AREABASE * ALTURA / 3)
Label5 = VOLUMENTOTAL

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Label5 = ""

End Sub

Private Sub Command3_Click()
End
End Sub

