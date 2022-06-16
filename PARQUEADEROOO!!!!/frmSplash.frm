VERSION 5.00
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4395
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7785
      Begin MARCHOSOLib.Marchoso M 
         Height          =   2175
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   3836
         _StockProps     =   1
         FileName        =   ""
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   7320
         Top             =   600
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
         Left            =   3720
         TabIndex        =   1
         Top             =   3270
         Width           =   3615
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
         Left            =   6000
         TabIndex        =   2
         Top             =   3600
         Width           =   1275
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
         Left            =   1710
         TabIndex        =   3
         Top             =   120
         Width           =   4485
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
M.FileName = "L:\PARQUEADEROOO!!!!\carro.gif"

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
Form1.Show
Timer1.Enabled = False
End If

End Sub

