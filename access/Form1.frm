VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   4560
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "inicio"
      Caption         =   "Entre registros"
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
   Begin VB.TextBox Text3 
      DataField       =   "Campo3"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Text            =   "Sexo"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "Campo4"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Text            =   "Carrera"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "Campo2"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Text            =   "Edad"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ultimo"
      Height          =   495
      Index           =   5
      Left            =   7560
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   495
      Index           =   4
      Left            =   7560
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      Height          =   495
      Index           =   3
      Left            =   7560
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   495
      Index           =   2
      Left            =   7560
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nombre "
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "Campo1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Text            =   "Nombre"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Cedula"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Sexo"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Edad"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
