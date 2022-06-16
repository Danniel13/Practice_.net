VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AHORCADO"
   ClientHeight    =   5655
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Registro 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton B_Comprobar 
      Caption         =   "Comprobar"
      Height          =   615
      Left            =   4200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Texto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   10
      Left            =   1800
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   9
      Left            =   1200
      Picture         =   "Form1.frx":1548
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   8
      Left            =   600
      Picture         =   "Form1.frx":267C
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   7
      Left            =   4800
      Picture         =   "Form1.frx":36EF
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   6
      Left            =   4200
      Picture         =   "Form1.frx":4692
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   5
      Left            =   3600
      Picture         =   "Form1.frx":551D
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   4
      Left            =   3000
      Picture         =   "Form1.frx":62AD
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   3
      Left            =   2400
      Picture         =   "Form1.frx":6E50
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   2
      Left            =   1800
      Picture         =   "Form1.frx":78B7
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":8262
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Horca 
      Height          =   495
      Index           =   0
      Left            =   600
      Picture         =   "Form1.frx":8AB7
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   5790
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   5310
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   4830
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   4350
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   3870
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   3390
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2910
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2430
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1950
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1470
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   990
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Linea 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   510
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Imagen 
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce letra"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   5760
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   5280
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4800
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Letra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
   Begin VB.Menu B_Juego 
      Caption         =   "Juego"
      Begin VB.Menu B_Nuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu B_Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim palabra As String, longitud As Integer, dibujito As Integer
Public Sub reiniciar()
'reinicia valores
For i = 0 To 11
Letra(i).Caption = ""
Linea(i).Visible = False
Next
Registro.Text = ""
Texto.Text = ""
dibujito = 0
Imagen.Picture = Horca(dibujito).Picture

palabra = LCase(Obtener_Palabra)
Debug.Print palabra
longitud = Len(palabra)
'Hace visibles lineas necesarias
For i = 1 To longitud
    Linea(i).Visible = True
    Letra(i).Visible = True
Next
End Sub
Public Sub escritor(word As String, palabra As String, dimension As Integer)
'determina la posicion que ocupa la letra en la palabra
'y hace visible el label correspondiente
For i = 1 To dimension
    If CStr(Mid(palabra, i, 1)) = word Then
        Letra(i).Caption = word
    End If
Next
End Sub
Public Function repeticiones(Letra As String, palabra As String, dimension As Integer) As Integer
'determina el numero de veces que se repite la letra
Dim contador As Integer
For i = 1 To dimension
    If CStr(Mid(palabra, i, 1)) = Letra Then contador = contador + 1
Next
repeticiones = contador
End Function
Public Function aleatorio(limiteinf As Integer, limitesup As Integer) As Integer
'Genera un número aleatorio
Randomize
aleatorio = Int((limitesup - limiteinf + 1) * Rnd + limiteinf)
End Function
Public Function ruta() As String
 'devuelve la ruta de la aplicación
    Dim path As String
    path = App.path
    If Right(path, 1) = "\" Then
        path = App.path
    Else
        path = App.path + "\"
    End If
    ruta = path
End Function
Public Function Obtener_Palabra() As String
Dim contador As Integer, palabra As String, numero As Integer
On Error GoTo fallo
Open ruta & "palabras.txt" For Input As #1
'cuenta el numero de palabras en el archivo
While Not EOF(1)
     Line Input #1, palabra
     contador = contador + 1
Wend
Close
'genera un numero entre 1 y el contador(numero de palabras)
numero = aleatorio(1, contador)
'obtiene palabra correspondiente a numero
Open ruta & "palabras.txt" For Input As #1
For i = 1 To numero
    Line Input #1, palabra
Next
Obtener_Palabra = palabra
Close: Exit Function
fallo:
MsgBox "No se ha encontrado el archivo palabras.txt", , "Ahorcado"
End
End Function

Private Sub B_Acerca_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub B_Comprobar_Click()
Dim letraentrada As String, win As Boolean
win = True
If Texto.Text = "" Then
    MsgBox "Introduce una letra", , "Ahorcado"
    Exit Sub
Else
    letraentrada = LCase(Texto.Text)
    Registro.Text = Registro.Text + letraentrada + " "
End If

If repeticiones(letraentrada, palabra, longitud) = 0 Then
    dibujito = dibujito + 1
    Imagen.Picture = Horca(dibujito).Picture
    Texto.Text = ""
Else
    escritor letraentrada, palabra, longitud
    Texto.Text = ""
End If

'Comprueba si se ha acertado la palabra
For i = 1 To longitud
    If Letra(i).Caption = "" Then win = False
Next
If win = True Then MsgBox "Bien!!! Acertaste la palabra", , "Ahorcado": reiniciar
'Comprueba si se ha perdido
If dibujito = 10 Then MsgBox "      GAME OVER" + vbCrLf + "La palabra era " + palabra, , "Ahorcado": reiniciar
End Sub

Private Sub B_Nuevo_Click()
reiniciar
End Sub

Private Sub B_Salir_Click()
Unload Form1
End Sub

Private Sub Form_Load()
reiniciar
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim mensaje1, titulo1, respuesta1, Botones1
Botones1 = vbYesNo + vbExclamation + vbDefaultButton2
mensaje1 = "¿Estás seguro de que quieres abandonar?"
titulo1 = "Ahorcado"
respuesta1 = MsgBox(mensaje1, Botones1, titulo1)

If respuesta1 = vbYes Then
End
Else
Cancel = 1
End If
End Sub

Private Sub Texto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then B_Comprobar.Value = True
End Sub

