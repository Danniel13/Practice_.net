VERSION 5.00
Begin VB.Form CALCULADORA 
   BackColor       =   &H00000000&
   Caption         =   "CALCULADORA"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "          CALCULADORA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6135
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   8295
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Secante"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cotangente"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox n 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Width           =   6615
      End
      Begin VB.CommandButton SUMA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4200
         Width           =   500
      End
      Begin VB.CommandButton RESTA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Width           =   500
      End
      Begin VB.CommandButton MULTIPLICACION 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3000
         Width           =   500
      End
      Begin VB.CommandButton DIVISION 
         BackColor       =   &H00C0C0C0&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton TOTAL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4800
         Width           =   495
      End
      Begin VB.CommandButton BORRAR 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Borrar Todo"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton RAIZ 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Raiz Cuadrada"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton SENO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SENO"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton EXPONENCIAL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exponencia"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton PORCENTAJE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CommandButton a1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton a2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton a3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton a4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton a5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton a6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton a7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton a8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton a9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton a0 
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   ","
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retroceso"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+/-"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton COSENO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COSENO"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton tangente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TANGENTE"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logaritmo"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pi"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4800
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cosecante"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3960
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "CALCULADORA"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1440
      Width           =   3435
   End
End
Attribute VB_Name = "CALCULADORA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opr As Integer
Dim aux1 As Double
Dim aux2 As Double

Private Sub a0_Click()
n.Text = n.Text + "0"
End Sub

Private Sub a1_Click()
n.Text = n.Text + "1"

End Sub

Private Sub a2_Click()
n.Text = n.Text + "2"
End Sub

Private Sub a3_Click()
n.Text = n.Text + "3"
End Sub

Private Sub a4_Click()
n.Text = n.Text + "4"
End Sub

Private Sub a5_Click()
n.Text = n.Text + "5"
End Sub

Private Sub a6_Click()
n.Text = n.Text + "6"
End Sub

Private Sub a7_Click()
n.Text = n.Text + "7"
End Sub

Private Sub a8_Click()
n.Text = n.Text + "8"
End Sub

Private Sub a9_Click()
n.Text = n.Text + "9"
End Sub

Private Sub BORRAR_Click()
n = ""
n.SetFocus
End Sub

Private Sub Command1_Click()
n.Text = n.Text + ","
End Sub

Private Sub Command2_Click()
n = Mid(n, 1, Len(n) - 1)
End Sub

Private Sub Command3_Click()
n.Text = n.Text + "-"
End Sub

Private Sub Command4_Click()

n = Log(n)

End Sub





Private Sub Command5_Click()
n = 3.1416
End Sub

Private Sub Command6_Click()
n = 1 / Cos(n * 3.1416 / 180)

End Sub

Private Sub Command7_Click()
n = 1 / Tan(n * 3.1416 / 180)
End Sub

Private Sub Command8_Click()
n = 1 / Sin(n * 3.1416 / 180)
End Sub

Private Sub COSENO_Click()
n = Cos(n * 3.1416 / 180)
End Sub

Private Sub DIVISION_Click()
aux1 = n
opr = 4
n = ""
n.SetFocus
End Sub

Private Sub EXPONENCIAL_Click()
n = Exp(n)
End Sub

Private Sub MULTIPLICACION_Click()
aux1 = n
opr = 3
n = ""
n.SetFocus
End Sub

Private Sub PORCENTAJE_Click()
n = n / 100

End Sub

Private Sub RAIZ_Click()
n = Sqr(n)
End Sub

Private Sub RESTA_Click()
aux1 = n
opr = 2
n = ""
n.SetFocus
End Sub

Private Sub SENO_Click()
n = Sin(n * 3.1416 / 180)
End Sub

Private Sub SUMA_Click()
aux1 = n
opr = 1
n = ""
n.SetFocus
End Sub

Private Sub tangente_Click()
n = Tan(n * 3.1416 / 180)
End Sub

Private Sub TOTAL_Click()
aux2 = n
If opr = 1 Then
n = aux1 + aux2
Else
If opr = 2 Then
n = aux1 - aux2
Else
If opr = 3 Then
n = aux1 * aux2
Else
If opr = 4 Then
n = aux1 / aux2

End If
End If
End If
End If
End Sub
