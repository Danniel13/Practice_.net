VERSION 5.00
Begin VB.Form frmCalculadora 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora Científica"
   ClientHeight    =   3525
   ClientLeft      =   4530
   ClientTop       =   2640
   ClientWidth     =   6180
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   18
      TabIndex        =   27
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Cubo 
      Caption         =   "X³"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Cuadrado 
      Caption         =   "X²"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Raiz 
      Caption         =   "Raiz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Tangente 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Coseno 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Seno 
      Caption         =   "Sen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Pi 
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2040
      TabIndex        =   11
      Top             =   720
      Width           =   1455
      Begin VB.CommandButton Porcentaje 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   840
         TabIndex        =   19
         Top             =   2040
         Width           =   480
      End
      Begin VB.CommandButton Operador 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   480
      End
      Begin VB.CommandButton Operador 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   1440
         Width           =   480
      End
      Begin VB.CommandButton Operador 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   480
      End
      Begin VB.CommandButton Operador 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   480
      End
      Begin VB.CommandButton Operador 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   480
      End
      Begin VB.CommandButton CE 
         Caption         =   "CE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton C 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Numero 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   840
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   480
   End
   Begin VB.CommandButton Numero 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      TabIndex        =   10
      Top             =   2640
      Width           =   480
   End
   Begin VB.Line Line6 
      X1              =   3120
      X2              =   3120
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Menu Cerrar 
      Caption         =   "Cerrar"
   End
   Begin VB.Menu Sobre 
      Caption         =   "Sobre..."
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Operador1, Operador2 As Double
Dim puntodecimal As Integer
Dim numerooperadores As Integer
Dim ultimatecla
Dim operacion
Dim temptext1


Private Sub C_Click(Index As Integer)
Text1 = Format(0, "0.")
    Operador1 = 0
    Operador2 = 0
    
End Sub

Private Sub CE_Click()
    Text1 = Format(0, "0.")
    puntodecimal = False
    ultimatecla = "CE"
End Sub

Private Sub Cerrar_Click()
End
End Sub

Private Sub Coseno_Click()
Dim rad2
rad2 = Text1 * (3.14059265358979 / 180)
Text1 = Sin(rad2)
End Sub

Private Sub Cuadrado_Click()
Text1 = Text1 ^ 2
End Sub

Private Sub Cubo_Click()
Text1 = Text1 ^ 3
End Sub

Private Sub Decimal_Click()
    If ultimatecla = "negativo" Then
        Text1 = Format(0, "-0.")
    ElseIf ultimatecla <> "numero" Then
        Text1 = Format(0, "0.")
    End If
    puntodecimal = True
    ultimatecla = "numero"
End Sub




Private Sub Form_Load()
    puntodecimal = False
    numerooperadores = 0
    ultimatecla = ""
    operacion = " "
    Text1 = Format(0, "0.")
 
End Sub




Private Sub Numero_Click(Index As Integer)

    If ultimatecla <> "numero" Then
        Text1 = Format(0, ".")
        puntodecimal = False
    End If
    If puntodecimal Then
        Text1 = Text1 + Numero(Index).Caption
    Else
        Text1 = Left(Text1, InStr(Text1, Format(0, ".")) - 1) + Numero(Index).Caption + Format(0, ".")
    End If
    If ultimatecla = "negativo" Then Text1 = "-" & Text1
    ultimatecla = "numero"
End Sub





Private Sub Pi_Click()
Dim Pi
Pi = 4 * Atn(1)
Text1 = Pi
End Sub

Private Sub Porcentaje_Click()
    Text1 = Text1 / 100
    ultimatecla = "operacion"
    operacion = "%"
    numerooperadores = numerooperadores + 1
    puntodecimal = True
End Sub


Private Sub Raiz_Click()
Text1 = Sqr(Text1)
End Sub

Private Sub Seno_Click()
Dim rad
rad = Text1 * (3.14059265358979 / 180)
Text1 = Sin(rad)
End Sub

Private Sub Sobre_Click()
frmSobre.Show
frmCalculadora.Enabled = False
End Sub


Private Sub Tangente_Click()
Dim rad3
rad3 = Text1 * (3.14059265358979 / 180)
Text1 = Sin(rad3)
End Sub






Private Sub Text1_Change()

End Sub
