VERSION 5.00
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form frmSplash1 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO"
   ClientHeight    =   4245
   ClientLeft      =   3990
   ClientTop       =   2970
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.TextBox Text2 
         BackColor       =   &H80000006&
         DataField       =   "Contraseña"
         DataSource      =   "Data1"
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         DataField       =   "Nombre"
         DataSource      =   "Data1"
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
         Height          =   495
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCELAR"
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ACEPTAR"
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
      End
      Begin MARCHOSOLib.Marchoso m 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   2990
         _StockProps     =   1
         BackColor       =   -2147483633
         FileName        =   ""
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Contraseña:"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
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
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1 = "usuario" And Text2 = "123" Then
Text1.Enabled = False And Text2.Visible = False
Text2.Enabled = False
frmSplash1.Hide
Form1.Show
Else
If Text1 = "duck" And Text2 = "platoon" Then
Text1.Enabled = False And Text2.Visible = False
Text2.Enabled = False
frmSplash1.Hide
Form1.Show

Else
MsgBox "Nombre de usuario y/o contraseña no validos favor de rectificarlos", 16, "Alerta"
End If
End If

End Sub




Private Sub Command2_Click()
End

End Sub



Private Sub Form_Load()
m.FileName = "F:\Registradora\user.gif"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()
End Sub
