VERSION 5.00
Begin VB.Form Registro 
   Caption         =   "Registro De Usuarios"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "L:\escolar\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir del Programa"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTRAR"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   2055
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2655
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
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Data1.UpdateRecord
Data1.Refresh
MsgBox "El Usuario ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
Registro.Hide
Alumnos.Show

End Sub

Private Sub Command2_Click()
End
End Sub
Private Sub Form_Activate()
Data1.Recordset.AddNew
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

