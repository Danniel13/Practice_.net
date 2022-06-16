VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   17010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
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
      Left            =   1920
      TabIndex        =   37
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8160
      TabIndex        =   31
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10080
      TabIndex        =   26
      Text            =   "Año"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "ESPECIALIZACION:"
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
      Height          =   2295
      Left            =   9360
      TabIndex        =   21
      Top             =   3360
      Width           =   3375
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000012&
         Caption         =   "Odontologia"
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
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000012&
         Caption         =   "Psiquiatria"
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
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000012&
         Caption         =   "Pediatria"
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
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
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
      Left            =   8400
      TabIndex        =   12
      Text            =   "Mes"
      Top             =   1320
      Width           =   1575
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
      Left            =   7320
      TabIndex        =   11
      Text            =   "Dìa"
      Top             =   1320
      Width           =   975
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
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
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
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   3015
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
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
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
         TabIndex        =   8
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "masculino"
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
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
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
      ItemData        =   "Form1.frx":0000
      Left            =   4800
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Top             =   7800
      Width           =   3255
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
      TabIndex        =   4
      Top             =   7800
      Width           =   4215
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
      TabIndex        =   3
      Top             =   10080
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
      TabIndex        =   2
      Top             =   10080
      Width           =   1515
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
      TabIndex        =   1
      Top             =   10080
      Width           =   1755
   End
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
      ItemData        =   "Form1.frx":0008
      Left            =   8520
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Top             =   7800
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Caption         =   "TIPO DE IDENTIFICACIÒN"
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
      Height          =   2175
      Left            =   840
      TabIndex        =   27
      Top             =   3240
      Width           =   3855
      Begin VB.OptionButton Option7 
         Caption         =   "Otra"
         Height          =   495
         Left            =   600
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Cedula de ciudadania"
         Height          =   495
         Left            =   600
         TabIndex        =   28
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "Ciudad:"
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
      Height          =   375
      Left            =   480
      TabIndex        =   36
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Left            =   7680
      TabIndex        =   35
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "Especializacion:"
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
      Height          =   495
      Left            =   5520
      TabIndex        =   34
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label10 
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
      Left            =   8160
      TabIndex        =   33
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "IDENTIFICACION:"
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
      Left            =   5640
      TabIndex        =   32
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Nùmero:"
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
      Left            =   6960
      TabIndex        =   30
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "HISTORIA CLINICA"
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
      Height          =   615
      Index           =   1
      Left            =   4320
      TabIndex        =   25
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "FECHA:"
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
      Left            =   6120
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
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
      Left            =   7680
      TabIndex        =   19
      Top             =   720
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
      Left            =   6120
      TabIndex        =   18
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Apellido:"
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
      Left            =   480
      TabIndex        =   17
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Nombre:"
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
      Left            =   480
      TabIndex        =   16
      Top             =   720
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   13080
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3375
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
      TabIndex        =   15
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FECHA"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ESPECIALIZACION"
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
      Left            =   8520
      TabIndex        =   13
      Top             =   7320
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Label12 = "Pediatria" Then
Image1.Picture = LoadPicture("E:\VISUAL 2012\parcial\ped.jpg")
End If

If Label12 = "Psiquiatria" Then
Image1.Picture = LoadPicture("E:\VISUAL 2012\parcial\psi.jpg")
End If


If Label12 = "Odontologia" Then
Image1.Picture = LoadPicture("E:\VISUAL 2012\parcial\odo.jpg")
End If



List1.AddItem Text1 & " " & " " & Text2
List1.AddItem "Sexo: " & Label5
List1.AddItem "Ciudad: " & Text4
List1.AddItem Label10
List1.AddItem "Nùmero: " & Text3
List2.AddItem Combo1 & " de " & Combo2 & " de " & Combo3

List3.AddItem Label12


End Sub


Private Sub Command3_Click()
Combo1.Text = ""
Label5 = ""
List1.Clear
Combo1.SetFocus

Combo3.Text = ""


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
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo1.AddItem "17"
Combo1.AddItem "18"
Combo1.AddItem "19"
Combo1.AddItem "20"
Combo1.AddItem "21"
Combo1.AddItem "22"
Combo1.AddItem "23"
Combo1.AddItem "24"
Combo1.AddItem "25"
Combo1.AddItem "26"
Combo1.AddItem "27"
Combo1.AddItem "28"
Combo1.AddItem "29"
Combo1.AddItem "30"
Combo1.AddItem "31"


Combo2.AddItem "Enero"
Combo2.AddItem "Febrero"
Combo2.AddItem "Marzo"
Combo2.AddItem "Abril"
Combo2.AddItem "Mayo"
Combo2.AddItem "Junio"
Combo2.AddItem "Julio"
Combo2.AddItem "Agosto"
Combo2.AddItem "Septiembre"
Combo2.AddItem "Octubre"
Combo2.AddItem "Noviembre"
Combo2.AddItem "Diciembre"


Combo3.AddItem "2000"
Combo3.AddItem "2001"
Combo3.AddItem "2002"
Combo3.AddItem "2003"
Combo3.AddItem "2004"
Combo3.AddItem "2005"
Combo3.AddItem "2006"
Combo3.AddItem "2007"
Combo3.AddItem "2008"
Combo3.AddItem "2009"
Combo3.AddItem "2010"
Combo3.AddItem "2011"
Combo3.AddItem "2012"
Combo3.AddItem "2013"



End Sub

Private Sub Option1_Click()
Label5 = Option1.Caption
End Sub

Private Sub Option2_Click()
Label5 = Option2.Caption
End Sub

Private Sub Option3_Click()
Label12 = Option3.Caption
End Sub

Private Sub Option4_Click()
Label12 = Option4.Caption
End Sub

Private Sub Option5_Click()
Label12 = Option5.Caption
End Sub

Private Sub Option6_Click()
Label10 = Option6.Caption

End Sub

Private Sub Option7_Click()
Label10 = Option7.Caption
End Sub
