VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form busqueda 
   Caption         =   "busqueda"
   ClientHeight    =   585
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "22:00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22/05/2012"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu opc 
      Caption         =   "Opciones"
      Begin VB.Menu est 
         Caption         =   "Estudiantes"
      End
      Begin VB.Menu bsq 
         Caption         =   "busquedas1"
      End
      Begin VB.Menu werwer 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asdas_Click()
busquedas2.Show

End Sub

Private Sub bsq_Click()
busquedas1.Show

End Sub

Private Sub est_Click()
Alumnos.Show

End Sub
