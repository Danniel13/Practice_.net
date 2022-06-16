VERSION 5.00
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BIENVEIDO AL SISTEMA"
   ClientHeight    =   4245
   ClientLeft      =   3780
   ClientTop       =   2970
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
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
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ingresar"
         Height          =   495
         Left            =   4320
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin MARCHOSOLib.Marchoso m 
         Height          =   2055
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   3625
         _StockProps     =   1
         BackColor       =   -2147483633
         FileName        =   ""
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "BIENVENIDO AL SISTEMA "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   450
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   4485
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00000000&
         Caption         =   "Compañía: The Duck Software"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   3390
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
frmSplash1.Show
frmSplash.Hide
End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
m.FileName = "F:\Registradora\image.gif"
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblLicenseTo_Click()

End Sub
