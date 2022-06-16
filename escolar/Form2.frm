VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Bienvenido al sistema escolar "
   ClientHeight    =   3090
   ClientLeft      =   4785
   ClientTop       =   3180
   ClientWidth     =   5730
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5730
   Begin VB.CommandButton Command2 
      Caption         =   "Salir del programa"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar al programa"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Introduce tu contraseña:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce tu nombre:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim usuario As String
Dim contrasena As String
Dim mensaje As String

Registro.Data1.Refresh

usuario = Text1
contrasena = Text2

Do Until Registro.Data1.Recordset.EOF
If Registro.Data1.Recordset.Fields("usuario").Value = usuario And Registro.Data1.Recordset.Fields("contrasena").Value = contrasena Then
MsgBox "Hola " & usuario & ", ¿Cómo has estado?", vbOKOnly, "Bienvenido al Programa"

Entrada.Hide
Alumnos.Show
Exit Sub

Else
Registro.Data1.Recordset.MoveNext
End If

Loop

mensaje = MsgBox("No te Conozco " & usuario & ", Intenta de Nuevo", vbOKOnly, "Atención, Usuario No Autorizado!!!")

If (mensaje = 1) Then
Entrada.Show
Text1 = ""
Text2 = ""

Else
End
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Form_activate()
Text1.SetFocus
End Sub

