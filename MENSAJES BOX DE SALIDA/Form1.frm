VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16500
   LinkTopic       =   "Form1"
   ScaleHeight     =   12210
   ScaleWidth      =   16500
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "MENSAJE 2"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MENSAJE 1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 64, "EJEMPLO1"
End Sub

Private Sub Command2_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS" & vbCrLf & "CALIFICACIONES", , "EJEMPLO 1"
End Sub

Private Sub Command3_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 16, "EJEMPLO"
End Sub

Private Sub Command4_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 32, "EJEMPLO"
End Sub

Private Sub Command5_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 48, "EJEMPLO"
End Sub

Private Sub Command6_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 64, "EJEMPLO"
End Sub

Private Sub Command7_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 0, "EJEMPLO"
End Sub

Private Sub Command8_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 2, "EJEMPLO"
End Sub

Private Sub Command9_Click()
MsgBox "BIENVENIDO AL SISTEMA DE NOTAS", 3, "EJEMPLO"
End Sub
