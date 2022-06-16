VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Form2Form2"
   ClientHeight    =   8415
   ClientLeft      =   1095
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton volver 
      Caption         =   "volver"
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton resultado 
      Caption         =   "ver resultados"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   6840
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0647
      OLEDBString     =   $"Form2.frx":06E3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tabla1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton bb 
      Caption         =   "votar"
      Height          =   495
      Left            =   12240
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cc 
      Caption         =   "votar"
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton oo 
      Caption         =   "votar"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton uu 
      Caption         =   "votar"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3150
      Left            =   120
      Picture         =   "Form2.frx":077F
      ScaleHeight     =   3090
      ScaleWidth      =   3030
      TabIndex        =   3
      Top             =   1560
      Width           =   3090
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   2925
      Left            =   3720
      Picture         =   "Form2.frx":34A4
      ScaleHeight     =   2865
      ScaleWidth      =   3210
      TabIndex        =   2
      Top             =   1560
      Width           =   3270
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   3105
      Left            =   7560
      Picture         =   "Form2.frx":5999
      ScaleHeight     =   3045
      ScaleWidth      =   3030
      TabIndex        =   1
      Top             =   1560
      Width           =   3090
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   3225
      Left            =   11400
      Picture         =   "Form2.frx":9F5B
      ScaleHeight     =   3165
      ScaleWidth      =   3570
      TabIndex        =   0
      Top             =   1440
      Width           =   3630
   End
   Begin VB.Label bl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "boto en blanco"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   13320
      TabIndex        =   19
      Top             =   5640
      Width           =   45
   End
   Begin VB.Label ch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "chavez"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9120
      TabIndex        =   18
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label ob 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "obama"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label ur 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "uribe"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   16
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label ganador 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5160
      TabIndex        =   14
      Top             =   7680
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "el ganador es:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   7920
      Width           =   1965
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de votos:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11040
      TabIndex        =   12
      Top             =   5640
      Width           =   1965
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de votos:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   5760
      Width           =   1965
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de votos:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   5760
      Width           =   1965
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de votos:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -240
      TabIndex        =   9
      Top             =   5760
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vote por Favor:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   3150
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub bb_Click()
Dim blan As String
Dim beb As String
bl = bl + 1
Adodc1.Recordset.Update
beb = MsgBox("usted a botado en blanco", vbInformation, "botacion ejecutada")
mbm = MsgBox("gracias por participar de la botacion", bvaceptar, "botacion satisfactoria")
Form1.Show
End Sub

Private Sub cc_Click()
Dim cha As String
Dim cvz As String
ch = ch + 1
Adodc1.Recordset.Update
cha = MsgBox("usted a botado por chavez", vbInformation, "botacion ejecutada")
cvz = MsgBox("gracias por participar de la botacion", bvaceptar, "botacion satisfactoria")
Form1.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
ur = 0
ob = 0
ch = 0
bl = 0
Form2.Refresh
uu.Enabled = True
oo.Enabled = True
cc.Enabled = True
bb.Enabled = True
resultado.Enabled = False
    volver.Enabled = False
End Sub

Private Sub oo_Click()
Dim oba As String
Dim mbm As String
ob = ob + 1
Adodc1.Recordset.Update
oba = MsgBox("usted a botado por obama", vbInformation, "botacion ejecutada")
mbm = MsgBox("gracias por participar de la botacion", bvaceptar, "botacion satisfactoria")
Form1.Show
End Sub

Private Sub resultado_Click()
If ur = 0 Then
If ch = 0 Then
If ob = 0 Then
If bl = 0 Then
ganador = "no han iniciado las botaciones"
End If
End If
End If
End If
If ur > ob Then
If ur > ch Then
If ur > bl Then
ganador = "uribe"
End If
End If
End If
    If ob > ur Then
    If ob > ch Then
    If ob > bl Then
        ganador = "obama"
    End If
End If
End If
        If ch > ur Then
        If ch > ob Then
        If ch > bl Then
            ganador = "chavez"
       End If
End If
End If
           If ur = ob Then
           If ur > ch Then
           If ur > bl Then
                ganador = "uribe y obama se iran a la segunda vuelta"
            End If
End If
End If
                If ur = ch Then
                If ur > ob Then
                If ur > bl Then
                    ganador = "uribe y chaves se iran a la segunda vueta"
                End If
End If
End If
                    If ur = bl Then
                    If ur > ob Then
                    If ur > ch Then
                        ganador = "uribe"
                    End If
End If
End If
                        If ob = ch Then
                        If ob > ur Then
                        If ob > bl Then
                         ganador = "obama y chavez se iran a la segunda vuelta"
                        End If
End If
End If
                            If ob = bl Then
                            If ob > ur Then
                            If ob > ch Then
                             ganador = "obama"
                             End If
End If
End If
If ch = bl Then
If ch > ur Then
If ch > ob Then
 ganador = "chavez"
End If
End If
End If
    If ob = ur Then
    If ch = ur Then
    If ur > bl Then
     ganador = "uribe, obama y chavez se iran a la segunda vuelta"
    End If
End If
End If
        If ur = ch Then
        If ur = bl Then
        If ur > ob Then
         ganador = "chavez y uribe se iran a la segunda vuelta"
        End If
End If
End If
            If ur = ob Then
            If ur = bl Then
            If ur > ch Then
                ganador = "uribe y obama se iran a la segunda vielta"
            End If
End If

                If ur = ob Then
                If ur = ch Then
                If ur = bl Then
                 ganador = "chavez, obama y uribe se iran a la segunda vuelta"
                 End If
    End If
        End If
                 If bl > ob Then
                If bl > ch Then
                If bl > ur Then
ganador = "los botos en blanco se le asignaran al candidato con mayor puntaje"
End If
    End If
        End If
             
                 

     
       End If
End Sub

Private Sub uu_Click()
Dim uri As String
Dim aub As String
ur = ur + 1
Adodc1.Recordset.Update
uri = MsgBox("usted a botado por uribe", vbInformation, "botacion ejecutada")
aub = MsgBox("gracias por participar de la botacion", bvaceptar, "botacion satisfactoria")
Form1.Show
End Sub

Private Sub volver_Click()
Form1.Show
Form2.Refresh
uu.Enabled = True
oo.Enabled = True
cc.Enabled = True
bb.Enabled = True
resultado.Enabled = False
    volver.Enabled = False
    ganador = ""
End Sub
