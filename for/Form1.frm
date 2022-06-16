VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "accedente"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "while"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "for accedente"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "for"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   6000
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = 1
For x = 10 To a Step -1
List1.AddItem x
Next x



End Sub

Private Sub Command2_Click()
a = 10
For x = 1 To a
List1.AddItem x
Next x

End Sub

Private Sub Command3_Click()
a = 1
While a < 10
List1.AddItem a
a = a + 1
Wend

End Sub

Private Sub Command4_Click()
a = 10
While a > 1
List1.AddItem a
a = a - 1
Wend

End Sub

Private Sub Command5_Click()
List1 = ""
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()

End Sub
