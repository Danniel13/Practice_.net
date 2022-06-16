VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   17160
   Begin VB.CommandButton Command2 
      Caption         =   "BORRAR"
      Height          =   615
      Left            =   5280
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "%"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "VALOR IVA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "VALOR ARTICULO:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "IVA DE UN PRODUCTO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim prec As Double
Dim iva As Double

End Sub

Private Sub Form_Load()

End Sub

Private Sub Label4_Click()

End Sub
