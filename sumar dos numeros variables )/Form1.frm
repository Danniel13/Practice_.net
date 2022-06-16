VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "suma de dos numeros"
   ClientHeight    =   13425
   ClientLeft      =   390
   ClientTop       =   -60
   ClientWidth     =   18450
   LinkTopic       =   "Form1"
   ScaleHeight     =   13425
   ScaleWidth      =   18450
   Begin VB.TextBox C 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox B 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox A 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FINALIZAR"
      Height          =   855
      Left            =   8040
      TabIndex        =   5
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   855
      Left            =   6480
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUMA"
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "SUMA DE DOS NUMEROS."
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   " RESULTADO"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "NUMERO 2"
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "NUMERO 1"
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
C = Val(A) + Val(B)

Text1.Text = " "
Text2.Text = " "
Text3.Text = " "





End Sub

Private Sub Label7_Click()

End Sub

