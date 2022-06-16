VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   12450
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   17160
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR"
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Valor de Pi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Introduzca el Radio:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "AREA DE UN CIRCULO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AREA As Double
Dim Pi As Double
Dim RADIO As Double
Pi = 3.14159
Label6 = Pi
RADIO = Val(Text1.Text)
AREA = Pi * (RADIO * RADIO)
Label7 = AREA
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Label6 = ""
Label7 = ""
End Sub

Private Sub Label4_Click()

End Sub

