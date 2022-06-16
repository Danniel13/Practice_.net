VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.ListBox List3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2160
      ItemData        =   "Form1.frx":0000
      Left            =   8280
      List            =   "Form1.frx":0002
      TabIndex        =   23
      Top             =   6120
      Width           =   2895
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
      Left            =   1560
      TabIndex        =   22
      Top             =   8400
      Width           =   1755
   End
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
      Left            =   5640
      TabIndex        =   21
      Top             =   8400
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Limpiar"
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
      Left            =   8760
      TabIndex        =   20
      Top             =   8400
      Width           =   1755
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   2160
      ItemData        =   "Form1.frx":0004
      Left            =   120
      List            =   "Form1.frx":0006
      TabIndex        =   18
      Top             =   6120
      Width           =   4215
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   2160
      ItemData        =   "Form1.frx":0008
      Left            =   4680
      List            =   "Form1.frx":000A
      TabIndex        =   17
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
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
      Left            =   1080
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "Macho"
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
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000012&
         Caption         =   "Hembra"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "HABITAD:"
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
      Left            =   8160
      TabIndex        =   10
      Top             =   2280
      Width           =   2895
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000012&
         Caption         =   "Tierra"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000012&
         Caption         =   "Agua"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000012&
         Caption         =   "Aire"
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
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
      Left            =   8640
      TabIndex        =   1
      Text            =   "Alimentación"
      Top             =   960
      Width           =   2295
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
      Left            =   8640
      TabIndex        =   0
      Text            =   "Respiración"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HABITAD"
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
      Index           =   3
      Left            =   8280
      TabIndex        =   27
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "RESPIRACION:"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   26
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "GENERALES:"
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
      Left            =   840
      TabIndex        =   25
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   11400
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Respiración:"
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
      Left            =   11880
      TabIndex        =   24
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "DATOS: "
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
      TabIndex        =   19
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Animal:"
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
      Height          =   720
      Left            =   600
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Especie:"
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
      Height          =   480
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Sexo:"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Alimentación"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Respiración"
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
      Index           =   0
      Left            =   5760
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo2.Text = "BRANQUEAL" Then
Image1.Picture = LoadPicture("E:\combinado\image1.jpg")
End If

If Combo2.Text = "PULMONAR" Then
Image1.Picture = LoadPicture("E:\combinado\image2.jpg")
End If


If Text1 = "" Then
MsgBox " DEBE INGRESAR EL NOMBRE DEL ANIMAL", 16, "BASE DE DATOS"
End If
If Check1.Value = 1 Then
List3.AddItem Check1.Caption
End If

If Check2.Value = 1 Then
List3.AddItem Check2.Caption
End If

If Check3.Value = 1 Then
List3.AddItem Check3.Caption
End If

List1.AddItem Text1 & " " & " " & Text2 & " " & Text3 & " " & Label5 & " " & Combo1
List2.AddItem Combo2
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Combo1.Text = ""
Label5 = ""
List1.Clear
Combo1.SetFocus

Combo2.Text = ""
Label5 = ""
List1.Clear
Combo2.SetFocus


Text1.Text = ""
Label5 = ""
List1.Clear
Text1.SetFocus

Text2.Text = ""
Label5 = ""
List2.Clear
Text2.SetFocus

Text2.Text = ""
Label5 = ""
List3.Clear
Text2.SetFocus

Image1 = Nothing

End Sub

Private Sub Form_Load()
Combo1.AddItem "CARNE"
Combo1.AddItem "VEGETALES"
Combo1.AddItem "CONCENTRADO"
Combo1.AddItem "VITAMINAS"

Combo2.AddItem "BRANQUEAL"
Combo2.AddItem "PULMONAR"
End Sub

Private Sub Option1_Click()
Label5 = Option1.Caption
End Sub

Private Sub Option2_Click()
Label5 = Option2.Caption
End Sub


