VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TANGENTE 
      Caption         =   "TAN"
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton COSENO 
      Caption         =   "COS"
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton PUNTO 
      Caption         =   "."
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton CERO 
      Caption         =   "0"
      Height          =   495
      Left            =   3840
      TabIndex        =   21
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton NUEVE 
      Caption         =   "9"
      Height          =   615
      Left            =   5880
      TabIndex        =   20
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton OCHO 
      Caption         =   "8"
      Height          =   615
      Left            =   4920
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton SIETE 
      Caption         =   "7"
      Height          =   615
      Index           =   1
      Left            =   3840
      TabIndex        =   18
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton SEIS 
      Caption         =   "6"
      Height          =   615
      Left            =   5880
      TabIndex        =   17
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton CINCO 
      Caption         =   "5"
      Height          =   615
      Left            =   4920
      TabIndex        =   16
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton CUATRO 
      Caption         =   "4"
      Height          =   615
      Index           =   0
      Left            =   3840
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton TRES 
      Caption         =   "3"
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton ODS 
      Caption         =   "2"
      Height          =   615
      Left            =   4920
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton UNO 
      Caption         =   "1"
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox N 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CommandButton SUMA 
      Caption         =   "+"
      Height          =   615
      Left            =   7200
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton RESTA 
      Caption         =   "-"
      Height          =   615
      Left            =   8160
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton MULTIPLICACION 
      Caption         =   "*"
      Height          =   615
      Left            =   7200
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton DIVISION 
      Caption         =   "/"
      Height          =   615
      Left            =   8160
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton TOTAL 
      Caption         =   "="
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton BORRAR 
      Caption         =   "AC"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton RAIZ 
      Caption         =   "RAIZ"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton SENO 
      Caption         =   "SEN"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton EXP 
      Caption         =   "EXP"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton PORCIENTO 
      Caption         =   "%"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CALCULADORA"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   3495
      Left            =   3600
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AUX1 As Integer
Dim AUX2 As Integer
Dim OPR As Integer

Private Sub BORRAR_Click()
N = ""
End Sub

Private Sub CERO_Click()
N.Text = N.Text & "0"
End Sub

Private Sub CINCO_Click()
N.Text = N.Text & "5"
End Sub

Private Sub COSENO_Click()
N = Cos(N * (3.14159 / 180))
End Sub

Private Sub CUATRO_Click(Index As Integer)
N.Text = N.Text & "4"
End Sub

Private Sub DIVISION_Click()
AUX1 = N
OPR = 4
N = ""
N.SetFocus
End Sub

Private Sub EXP_Click()

End Sub

Private Sub MULTIPLICACION_Click()
AUX1 = N
OPR = 3
N = ""
N.SetFocus

End Sub

Private Sub NUEVE_Click()
N.Text = N.Text & "9"
End Sub

Private Sub OCHO_Click()
N.Text = N.Text & "8"
End Sub

Private Sub ODS_Click()
N.Text = N.Text & "2"
End Sub

Private Sub PORCIENTO_Click()
AUX1 = N
OPR = 5
N = ""
N.SetFocus

End Sub

Private Sub RAIZ_Click()
N = Sqr(N)
End Sub

Private Sub RESTA_Click()
AUX1 = N
OPR = 2
N = ""
N.SetFocus

End Sub

Private Sub SEIS_Click()
N.Text = N.Text & "6"
End Sub

Private Sub SENO_Click()
N = Sin(N * (3.14159 / 180))
End Sub

Private Sub SIETE_Click(Index As Integer)
N.Text = N.Text & "7"
End Sub

Private Sub SUMA_Click()
AUX1 = N
OPR = 1
N = ""
N.SetFocus

End Sub


Private Sub TANGENTE_Click()
N = Tan(N * (3.14159 / 180))
End Sub

Private Sub TOTAL_Click()
AUX2 = Val(N)
If OPR = 1 Then
N = AUX1 + AUX2
Else
If OPR = 2 Then
N = AUX1 - AUX2
Else
If OPR = 3 Then
N = AUX1 * AUX2
Else
If OPR = 4 Then
N = AUX1 / AUX2
Else
If OPR = 5 Then
N = AUX1 * AUX2 / 100
End If
End If
End If
End If
End If
End Sub

Private Sub TRES_Click()
N.Text = N.Text & "3"
End Sub

Private Sub UNO_Click()
N.Text = N.Text & "1"
End Sub
