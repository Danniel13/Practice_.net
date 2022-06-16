VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   12240
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   12240
   ScaleWidth      =   17160
   Begin VB.CommandButton Command1 
      Caption         =   "DATOS"
      Height          =   735
      Left            =   5280
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "CARGO"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "NOMBRES"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "APELLIDOS"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim mensaje As String, nombre As String, apellido As String, cargo As String

mensaje = "escriba su nombre"
nombre = InputBox$(mensaje)


mensaje = "escriba su apellido"
apellido = InputBox$(mensaje)


mensaje = "escriba su cargo"
cargo = InputBox$(mensaje)

Label4.Caption = nombre
Label5.Caption = apellido
Label6.Caption = cargo
End Sub

