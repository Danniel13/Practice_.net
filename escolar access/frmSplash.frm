VERSION 5.00
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   4185
   ClientTop       =   3165
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   7200
         Top             =   120
      End
      Begin MARCHOSOLib.Marchoso M 
         Height          =   3015
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483641
         FileName        =   ""
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
         Left            =   6120
         TabIndex        =   5
         Top             =   4200
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
         Index           =   1
         Left            =   3960
         TabIndex        =   4
         Top             =   3840
         Width           =   3615
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000012&
         Caption         =   "Sistema Escolar."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "BIENVENIDO "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   675
         Left            =   4080
         TabIndex        =   1
         Top             =   1080
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_load()
M.FileName = "F:\escolar\image.gif"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Static J As Integer
J = J + 1
If J = 20 Then
frmSplash.Hide
Bienvenida.Show
Timer1.Enabled = False
End If

End Sub
