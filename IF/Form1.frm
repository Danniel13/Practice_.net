VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12900
   ClientLeft      =   420
   ClientTop       =   630
   ClientWidth     =   20625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12900
   ScaleWidth      =   20625
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO"
      Height          =   735
      Left            =   12840
      TabIndex        =   4
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   12360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   12360
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "CONTRASEÑA"
      Height          =   615
      Left            =   9360
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      Height          =   735
      Left            =   9360
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "duck" And Text2 = "secreto" Then
Form2.Show
Else
MsgBox ("clave erronea")
End If
Text1 = ""
Text2 = ""

End Sub

Private Sub Form_Load()

End Sub
