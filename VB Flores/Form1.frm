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
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   855
      Left            =   1200
      TabIndex        =   14
      Top             =   9720
      Width           =   2295
   End
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
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6000
      TabIndex        =   12
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000FF&
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
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
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
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
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
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "FLOR 3"
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
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "FLOR 2"
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
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "FLOR 1"
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   2055
      Left            =   2520
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "FOTO"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   8760
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "INFORMACION"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Nombre científico :"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   6240
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Nombre Vulgar O latino:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "ALGUNAS FLORES DEL MUNDO"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3240
      TabIndex        =   6
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
Text1.Text = "Amapola"
Text2.Text = "Papaver rhoeas"
Text3.Text = "La amapola es una hierba anual que alcanza a medir hasta 70 centímetros de altura y más. El tallo es erecto y lo recubren uno finos pelitos. Sus hojas son dentadas, alargadas y lobuladas, no poseen pecíolo, brotan alternas a lo largo del tallo. Las flores son muy llamativas, poseen delicados pétalos luciendo un bello color rojo naranja, aunque existen otras de rojo intenso, amarillo, violeta y blanco."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\VB Flores\ama.jpg")
End Sub

Private Sub Command2_Click()
Text1.Text = "Azucena"
Text2.Text = "Lilium"
Text3.Text = "La azucena comprende más de 80 especies. Son plantas bulbosas que llegan a medir 1 metro de altura, su tallo es rígido y al final sostiene entre 8 y 12 flores con forma de trompeta que nacen de un mismo lugar. Los pétalos de estas flores son curvados hacia atrás. Las azucenas desprenden un exquisito y suave perfume, sobre todo durante la noche, pero no todas las variedades poseen aroma."
Image1.Picture = LoadPicture("C:\Documents and Settings\Administrador\Escritorio\VB Flores\azu.jpg")
End Sub

Private Sub Command3_Click()
Text1.Text = "Caléndula"
Text2.Text = "Calendula officinalis"
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

