VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   4215
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   7320
      TabIndex        =   13
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SIGUIENTE"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allan Poe"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Miguel de Cervantes"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pablo Neruda"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Gabriela Mistral"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "Gabriel Garcia Marquez"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   1440
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   2055
      Left            =   3720
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "FOTOGRAFIA"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8520
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "BIOGRAFIA"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Ciudad de Origen"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   7320
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Obra Reconocida"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   4200
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "ALGUNOS ESCRITORES FAMOSOS"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text1.Text = "Cien Años de Soledad."
Text2.Text = "Colombia."
Text3.Text = "Gabriel García Márquez nació en Aracataca (Magdalena), el 6 de marzo de 1927. Creció como niño único entre sus abuelos maternos y sus tías, pues sus padres, el telegrafista Gabriel Eligio García y Luisa Santiaga Márquez, se fueron a vivir, cuando Gabriel sólo contaba con cinco años, a la población de Sucre, donde don Gabriel Eligio montó una farmacia y donde tuvieron a la mayoría de sus once hijos."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\duck!!!\visual basic\nuevo\gabo.jpg")
End Sub

Private Sub Command2_Click()
Text1.Text = "Desolación."
Text2.Text = "Chile."
Text3.Text = "Lucila de María del Perpetuo Socorro Godoy Alcayaga, conocida por su seudónimo Gabriela Mistral (Vicuña, 7 de abril de 1889 – Nueva York, 10 de enero de 1957), fue una destacada poetisa, diplomática, feminista,1 y pedagoga chilena. Gabriela Mistral, una de las principales figuras de la literatura chilena y latinoamericana, fue la primera persona de América Latina en ganar el Premio Nobel de Literatura,2 el cual recibió en 1945."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\duck!!!\visual basic\nuevo\mist.jpg")
End Sub

Private Sub Command3_Click()
Text1.Text = "20 Poemas de Amor y una cancion desesperada."
Text2.Text = "Chile."
Text3.Text = "Ricardo Eliecer Neftalí Reyes Basoalto, Pablo Neruda (Parral, 12 de julio de 1904 - Santiago, 23 de septiembre de 1973), fue un poeta y militante comunista chileno, considerado entre los mejores y más influyentes artistas de su siglo, siendo llamado por el novelista Gabriel García Márquez «el más grande poeta del siglo XX en cualquier idioma»."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\duck!!!\visual basic\nuevo\ner.jpg")
End Sub

Private Sub Command4_Click()
Text1.Text = "Don Quijote de La mancha."
Text2.Text = "España."
Text3.Text = "Miguel de Cervantes Saavedra (¿Alcalá de Henares?, 29 de septiembre de 1547 – Madrid, 22 de abril[1] de 1616) fue un soldado, novelista, poeta y dramaturgo español. Es considerado una de las máximas figuras de la literatura española y universalmente conocido por haber escrito Don Quijote de la Mancha."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\duck!!!\visual basic\nuevo\cerv.jpg")
End Sub

Private Sub Command5_Click()
Text1.Text = " La narración de Arthur Gordon Pym."
Text2.Text = "Boston."
Text3.Text = "Edgar Allan Poe (Boston, Estados Unidos, 19 de enero de 1809 – Baltimore, Estados Unidos, 7 de octubre de 1849) fue un escritor, poeta, crítico y periodista romántico1 estadounidense, generalmente reconocido como uno de los maestros universales del relato corto, del cual fue uno de los primeros practicantes en su país. Fue renovador de la novela gótica, recordado especialmente por sus cuentos de terror."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\duck!!!\visual basic\nuevo\poe.jpg")
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Command7_Click()
End
End Sub
