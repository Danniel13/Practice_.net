VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "BASE DE DATOS"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1920
      Top             =   4440
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USECHE\Documents\DATOS\INICIO.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USECHE\Documents\DATOS\INICIO.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "1"
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
   Begin VB.CommandButton ULTIMO 
      Caption         =   "ULTIMO REGISTRO"
      Height          =   435
      Left            =   7920
      TabIndex        =   13
      Top             =   4680
      Width           =   2100
   End
   Begin VB.CommandButton ANTERIOR 
      Caption         =   "REGISTRO ANTERIOR"
      Height          =   435
      Left            =   7920
      TabIndex        =   12
      Top             =   3960
      Width           =   2100
   End
   Begin VB.CommandButton SIGUIENTE 
      Caption         =   "SIGUIENTE REGISTRO"
      Height          =   435
      Left            =   7920
      TabIndex        =   11
      Top             =   3240
      Width           =   2100
   End
   Begin VB.CommandButton ELIMINAR 
      Caption         =   "ELIMINAR"
      Height          =   435
      Left            =   7920
      TabIndex        =   10
      Top             =   2400
      Width           =   2100
   End
   Begin VB.CommandButton GUARDAR 
      Caption         =   "GUARDAR REGISTRO"
      Height          =   435
      Left            =   7920
      TabIndex        =   9
      Top             =   1560
      Width           =   2100
   End
   Begin VB.CommandButton NUEVO 
      Caption         =   "NUEVO REGISTRO"
      Height          =   435
      Left            =   7920
      TabIndex        =   8
      Top             =   720
      Width           =   2100
   End
   Begin VB.TextBox CARRERA 
      DataField       =   "CARRERE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3000
      Width           =   2170
   End
   Begin VB.TextBox SEXO 
      DataField       =   "SEXO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2280
      Width           =   2170
   End
   Begin VB.TextBox EDAD 
      DataField       =   "EDAD"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   2170
   End
   Begin VB.TextBox NOMBRE 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   2170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CARRERA"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   3120
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SEXO"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "EDAD"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE"
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ANTERIOR_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub ELIMINAR_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub GUARDAR_Click()
Adodc1.Recordset.Update
End Sub

Private Sub NUEVO_Click()
Adodc1.Recordset.AddNew
End Sub


Private Sub SIGUIENTE_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub ULTIMO_Click()
Adodc1.Recordset.MoveLast
End Sub
