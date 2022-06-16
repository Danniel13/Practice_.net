VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   12450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   13112.16
   ScaleMode       =   0  'User
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   7560
      TabIndex        =   24
      Top             =   11520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "Listar"
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
      Left            =   4560
      TabIndex        =   23
      Top             =   11520
      Width           =   2535
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      ItemData        =   "Form1.frx":0000
      Left            =   9360
      List            =   "Form1.frx":0002
      TabIndex        =   22
      Top             =   8520
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8160
      TabIndex        =   21
      Text            =   "Profesion"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8160
      TabIndex        =   20
      Text            =   "Estado Civil"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      ItemData        =   "Form1.frx":0004
      Left            =   4920
      List            =   "Form1.frx":0006
      TabIndex        =   19
      Top             =   8520
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      ItemData        =   "Form1.frx":0008
      Left            =   960
      List            =   "Form1.frx":000A
      TabIndex        =   18
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "ESTUDIOS"
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
      Height          =   4215
      Left            =   8280
      TabIndex        =   13
      Top             =   3000
      Width           =   2895
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000012&
         Caption         =   "Especializado"
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
         Left            =   360
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000012&
         Caption         =   "Profesional"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000012&
         Caption         =   "Técnico"
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
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "SEXO:"
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
      Height          =   2415
      Left            =   1920
      TabIndex        =   10
      Top             =   3120
      Width           =   3615
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000012&
         Caption         =   "Femenino"
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
         Left            =   480
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "Maculino"
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
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "@Kozuka Mincho Pro R"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "DATOS GENERALES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   4080
      TabIndex        =   17
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Profesión"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Estado civil"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
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
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Género"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Apellidos"
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
      Height          =   475
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Nombres"
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
      Height          =   475
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   475
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
List3.AddItem Check1.Caption
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
List3.AddItem Check2.Caption
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
List3.AddItem Check3.Caption
End If
End Sub

Private Sub Command1_Click()
If Text1 = "" Then
MsgBox " DEBE INGRESAR EL NUMERO DE LA CEDULA", 16, "BASE DE DATOS"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub List1_Click()
List1.AddItem Text1 & " " & " " & Text2 & " " & Text3 & " " & Label5 & " " & Combo1
List2.AddItem Combo2
End Sub

Private Sub Option1_Click()
Label5 = Option1.Caption
End Sub

Private Sub Option2_Click()
Label5 = Option2.Caption
End Sub
