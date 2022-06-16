VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton resultado 
      Caption         =   "visualizar resultados hasta el momento"
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   3765
      HideSelection   =   0   'False
      Left            =   10560
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Form1.frx":0647
      Top             =   2280
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   6360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
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
      Connect         =   $"Form1.frx":0749
      OLEDBString     =   $"Form1.frx":07E5
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
   Begin VB.CommandButton ingresar 
      Caption         =   "ingresar"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA:"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10440
      TabIndex        =   9
      Top             =   2160
      Width           =   60
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡¡adevertencia!!"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11280
      TabIndex        =   8
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      DataField       =   "cedula"
      DataSource      =   "Adodc1"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      DataField       =   "Id"
      DataSource      =   "Adodc1"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "participante numero"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   2550
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   4
      Top             =   7560
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "recurde que su sedula deve estar pre inscrita para poder botar"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   7575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cedula del participante"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   1680
      TabIndex        =   1
      Top             =   4320
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "biemvenido"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   5400
      TabIndex        =   0
      Top             =   1200
      Width           =   2790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ingresar_Click()
Dim dato As String
dato = InputBox("ingrese la cedula de el usuario", "buscar por registro")
campo = "cedula" & dato & "'" ' la variable es el nombre del campo.

Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "cedula=" & dato
 a = 1


     
If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
MsgBox "EL REGISTRO NO EXISTE", vbCritical ', "ERROR"
a = 2
End If


If a = 2 Then
     Form1.Show
     Else
     If a = 1 Then
     Form2.Show
    End If
    End If

     
    
End Sub

Private Sub resultado_Click()
Form2.Show
Form2.Refresh
Form2.uu.Enabled = False
Form2.Refresh
Form2.oo.Enabled = False
Form2.cc.Enabled = False
Form2.bb.Enabled = False
Form2.resultado.Enabled = True
    Form2.volver.Enabled = True

End Sub
