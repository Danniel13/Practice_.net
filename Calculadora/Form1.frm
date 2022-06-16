VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CALCULADORA"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LIMPIAR 
      Caption         =   "LIMPIAR"
      Height          =   615
      Left            =   7680
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR"
      Height          =   615
      Left            =   5400
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   11040
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton Option4 
         Caption         =   "DIVIDIR"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   2760
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         Caption         =   "MULTIPLICAR"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "RESTAR"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SUMAR"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   5640
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "RESULTADO"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "NUMERO 2"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "NUMERO 1"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A As Long
Dim B As Long
A = Val(Text1.Text)
B = Val(Text2.Text)

If Option1 Then
Label4 = A + B
End If
If Option2 Then
Label4 = A - B
End If
If Option3 Then
Label4 = A * B
End If
If Option4 Then
Label4 = A / B
End If

End Sub

