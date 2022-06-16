VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form plac 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Listfecha 
      BackColor       =   &H0080C0FF&
      Height          =   1425
      ItemData        =   "Form1.frx":0000
      Left            =   9120
      List            =   "Form1.frx":0002
      TabIndex        =   23
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton ingresar 
      BackColor       =   &H00000080&
      Caption         =   "ingresar"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox cuantos 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   8880
      TabIndex        =   16
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton salir 
      BackColor       =   &H00000080&
      Caption         =   "salir"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9120
      Width           =   1575
   End
   Begin VB.ListBox Listentra 
      BackColor       =   &H0080C0FF&
      Height          =   1425
      ItemData        =   "Form1.frx":0004
      Left            =   6840
      List            =   "Form1.frx":000B
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.ListBox listplaca 
      BackColor       =   &H0080C0FF&
      Height          =   1425
      ItemData        =   "Form1.frx":001A
      Left            =   3240
      List            =   "Form1.frx":0021
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6240
      Width           =   2775
   End
   Begin VB.ListBox listcedula 
      BackColor       =   &H0080C0FF&
      Height          =   1425
      ItemData        =   "Form1.frx":0030
      Left            =   120
      List            =   "Form1.frx":0037
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   6240
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8400
      Top             =   840
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"Form1.frx":0047
      OLEDBString     =   $"Form1.frx":00E8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9840
      Top             =   3240
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fecha y hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7800
      TabIndex        =   25
      Top             =   2400
      Width           =   2130
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fecha de hoy"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8400
      TabIndex        =   24
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label hsalida 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label hentrada 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label cedula 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label placa 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label entra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "h entrada"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6240
      TabIndex        =   14
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "placa"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4200
      TabIndex        =   13
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label cedul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cedula"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   10
      Top             =   5520
      Width           =   1140
   End
   Begin VB.Label time2 
      BackColor       =   &H0080C0FF&
      Caption         =   "00"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label time1 
      BackColor       =   &H0080C0FF&
      Caption         =   "00"
      Height          =   615
      Left            =   7680
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label saldo 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5415
      TabIndex        =   6
      Top             =   4080
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "saldo total"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hora de salida"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hora de entrada"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de la cedula de el usuario"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   4725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero de la placa"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000040C0&
      Caption         =   "programa pra el funcionamiento de un parqueadero"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   11070
   End
End
Attribute VB_Name = "plac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Label7.Caption = "00"
Label8.Caption = "00"
Label9.Caption = "00"
End Sub

Private Sub listusuario_Click()
listusuario.addltem
End Sub

Private Sub ingresar_Click()
Dim na As String
Dim nb As String
Dim nc As String
Dim nd As String


na = InputBox("ingrese el numero de su placa", "placas")
listplaca.AddItem na
placa = na


nb = InputBox("ingrese su numero de cedula", "cedilas")
listcedula.AddItem nb
cedula = nb


nc = time1
Listentra.AddItem nc
hentrada = nc

nd = time2
Listfecha.AddItem nd
hsalida = nd



End Sub

Private Sub orden_Click()

End Sub

Private Sub salir_Click()
End
End Sub

Private Sub Timer1_Timer()
time1.Caption = Time
time2.Caption = Date



End Sub

Private Sub Timer2_Timer()
Label10.Caption = Val(Label10.Caption) + 1
End Sub
