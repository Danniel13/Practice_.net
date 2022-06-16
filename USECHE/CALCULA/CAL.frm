VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton porcentage 
      Caption         =   " %"
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton exponencial 
      Caption         =   "X^y"
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton seno 
      Caption         =   "SENO"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton raiz 
      Caption         =   "RAIZ"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton borrar 
      Caption         =   "AC"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton igual 
      Caption         =   "="
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton dividir 
      Caption         =   "/"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   3720
      Width           =   490
   End
   Begin VB.CommandButton multiplicar 
      Caption         =   "*"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   3720
      Width           =   490
   End
   Begin VB.CommandButton resta 
      Caption         =   "-"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3720
      Width           =   490
   End
   Begin VB.CommandButton suma 
      Caption         =   "+"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3720
      Width           =   490
   End
   Begin VB.TextBox N 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CALCULADORA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   0
      Top             =   1560
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aux1 As Double
Dim aux2 As Double
Dim opr As Double

Private Sub igual_Click()
aux2 = Val(N)
If opr = 1 Then
N = aux1 + aux2
End If
End Sub

Private Sub raiz_Click()
N = Sqr(N)
End Sub

Private Sub seno_Click()
N = Sin(N)
End Sub

Private Sub suma_Click()
aux1 = N
opr = 1
N = ""
N.SetFocus

End Sub
