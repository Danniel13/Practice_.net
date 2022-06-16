VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   540
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   16155
   Begin VB.PictureBox p1 
      Height          =   3135
      Left            =   8880
      ScaleHeight     =   3075
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton finalizar 
      Caption         =   "FINALIZAR"
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
      Left            =   5880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton limpiar 
      Caption         =   "LIMPIAR"
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
      Left            =   3600
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton multiplicar 
      Caption         =   "MULTIPLICAR"
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
      TabIndex        =   3
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox n 
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
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Ingrese Un Nùmero"
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
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label as 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "TABLAS DE MULTIPLICAR"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   4305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub multiplicar_Click()
If n = "" Then
MsgBox " no hay numero"
Else
For i = 1 To 12
m = i * Val(n)
p1.Print i; "*"; n; "="; m
Next i
End If


End Sub
