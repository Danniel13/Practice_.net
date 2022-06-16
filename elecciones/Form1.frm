VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   315
   ClientTop       =   645
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "EL GANADOR ES:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   8400
      TabIndex        =   24
      Top             =   6240
      Width           =   3495
      Begin VB.Label e 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.TextBox a 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   4300
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   11520
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   8280
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4300
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VOTAR"
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
      Left            =   11520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VOTAR"
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
      Left            =   8280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "VOTAR"
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
      Left            =   5040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VOTAR"
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
      Left            =   1320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Total Votos:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5175
      Index           =   1
      Left            =   3480
      TabIndex        =   12
      Top             =   5160
      Width           =   3735
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   615
         Left            =   2040
         TabIndex        =   23
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox d 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox c 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   465
         Left            =   2040
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox b 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         Left            =   2040
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "MOSTRAR RESULTADOS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         MaskColor       =   &H000000FF&
         TabIndex        =   13
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "TOTAL VOTOS:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "BLANCO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "PETRO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "MOCKUS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "SANTOS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label m 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PETRO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   2235
      Left            =   7560
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "BLANCO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11040
      TabIndex        =   2
      Top             =   3120
      Width           =   2505
   End
   Begin VB.Label n 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "MOCKUS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   500
      Left            =   4560
      TabIndex        =   1
      Top             =   3000
      Width           =   2500
   End
   Begin VB.Label v 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "SANTOS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   500
      Left            =   840
      TabIndex        =   0
      Top             =   3000
      Width           =   2500
   End
   Begin VB.Image Image3 
      Height          =   2145
      Left            =   11280
      Picture         =   "Form1.frx":99012
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2280
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   4200
      Picture         =   "Form1.frx":A80CC
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   600
      Picture         =   "Form1.frx":BF55E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim santos  As Integer
 Dim mockus As Integer
 Dim petro As Integer
 Dim blanco As Integer
 
Private Sub Command1_Click()
   
Print santos
santos = santos + 1

Text5.Text = santos

MsgBox "SU VOTO A SIDO CONTADO.", 64, "Información"

End Sub

Private Sub Command2_Click()
Print mockus
mockus = mockus + 1
Text6.Text = mockus
MsgBox "SU VOTO A SIDO CONTADO.", 64, "Información"


End Sub

Private Sub Command3_Click()
Print petro
petro = petro + 1
Text7.Text = petro
MsgBox "SU VOTO A SIDO CONTADO.", 64, "Información"

End Sub

Private Sub Command4_Click()
Print blanco
blanco = blanco + 1
Text8.Text = blanco
MsgBox "SU VOTO A SIDO CONTADO.", 64, "Información"

End Sub

Private Sub Command5_Click()



a = Text5.Text
b = Text6.Text
c = Text7.Text
d = Text8.Text
text1 = Val(a) + Val(b) + Val(c) + Val(d)




Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "

Print
santos = 0
petro = 0
mockus = 0
blanco = 0

If Val(a) > Val(b) And Val(a) > Val(c) And Val(a) > Val(d) Then
MsgBox "Juan Manuel Santos", 64, "EL GANADOR ES: "
e = v
Else
If Val(b) > Val(a) And Val(b) > Val(c) And Val(b) > Val(d) Then
MsgBox "Antanas Mockus", 64, "EL GANADOR ES: "
e = n
Else
If Val(c) > Val(a) And Val(c) > Val(b) And Val(c) > Val(d) Then
MsgBox "Gustavo Petro", 64, "EL GANADOR ES: "
e = m
Else

If Val(d) > Val(a) And Val(d) > Val(b) And Val(d) > Val(c) Then
MsgBox "Ninguno, elección anulada ", 64, "EL GANADOR ES: "
e = "Ninguno"
Else

If Val(a) = Val(b) And Val(a) = Val(c) And Val(a) = Val(d) Then
MsgBox "Ninguno, es un empate C':", 64, "EL GANADOR ES: "
e = "Ninguno :)"
Else

If Val(b) = Val(a) And Val(b) = Val(c) And Val(b) = Val(d) Then
MsgBox "Ninguno, es un empate C':", 64, "EL GANADOR ES: "
e = n
Else
If Val(c) = Val(a) And Val(c) = Val(b) And Val(c) = Val(d) Then
MsgBox "Ninguno, es un empate C':", 64, "EL GANADOR ES: "
e = " Ninguno :) "
Else
If Val(d) = Val(a) And Val(d) = Val(b) And Val(d) = Val(c) Then
MsgBox "Ninguno, es un empate C': ", 64, "EL GANADOR ES: "
e = "Ninguno :)"

End If
End If
End If
End If
End If
End If
End If

End If



End Sub

Private Sub Image5_Click()

End Sub

Private Sub Text1_Change()
text1.Locked = True

End Sub

Private Sub Text2_Change()
Text2.Locked = True

End Sub

Private Sub Text3_Change()
Text3.Locked = True

End Sub

Private Sub Text4_Change()
Text4.Locked = True

End Sub

Private Sub Text5_Change()
Text5.Locked = True

End Sub

Private Sub Text6_Change()
Text6.Locked = True
End Sub

Private Sub Text7_Change()
Text7.Locked = True

End Sub

Private Sub Text8_Change()
Text8.Locked = True

End Sub
