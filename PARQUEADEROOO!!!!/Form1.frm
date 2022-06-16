VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Parqueadero"
   ClientHeight    =   9465
   ClientLeft      =   450
   ClientTop       =   1140
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   13995
   Begin VB.CommandButton Command7 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9120
      TabIndex        =   33
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ultimo Registro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   32
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Primer Registro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   31
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Anterior registro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   30
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente Registro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   29
      Top             =   6960
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4440
      Top             =   8880
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
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
      OLEDBString     =   $"Form1.frx":016A
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
   Begin VB.CommandButton Command2 
      Caption         =   "Total "
      Height          =   735
      Left            =   1680
      TabIndex        =   28
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parar Tiempo"
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9600
      Top             =   240
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      DataField       =   "# Matrícula"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   475
      Left            =   1560
      TabIndex        =   17
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      DataField       =   "# Identificación"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Caption         =   "TIPO DE IDENTIFICACIÒN"
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
      Height          =   2175
      Left            =   6000
      TabIndex        =   8
      Top             =   3600
      Width           =   3855
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000008&
         Caption         =   "Cedula de ciudadania"
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
         Height          =   495
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H80000012&
         Caption         =   "Otra"
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
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   480
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   475
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6120
      TabIndex        =   2
      Text            =   "Dìa"
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7200
      TabIndex        =   1
      Text            =   "Mes"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8880
      TabIndex        =   0
      Text            =   "Año"
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000012&
      DataField       =   "Total en Pesos:"
      DataSource      =   "Adodc1"
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
      Height          =   735
      Left            =   7080
      TabIndex        =   35
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      Caption         =   "TOTAL EN PESOS: "
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
      Height          =   495
      Left            =   4560
      TabIndex        =   34
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000012&
      DataField       =   "Tiempo en minutos"
      DataSource      =   "Adodc1"
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
      Height          =   495
      Left            =   2640
      TabIndex        =   27
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "TOTAL TIEMPO:"
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
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      DataField       =   "Hora Salida"
      DataSource      =   "Adodc1"
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
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "HORA SALIDA:"
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
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      DataField       =   "Fecha:"
      DataSource      =   "Adodc1"
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
      Height          =   495
      Left            =   4920
      TabIndex        =   22
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "FECHA DEL SISTEMA:"
      Enabled         =   0   'False
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
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "HORA DEL SISTEMA:"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
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
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      DataField       =   "Hora Ingreso"
      DataSource      =   "Adodc1"
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
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Número de Matrícula:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "HORA INGRESO:"
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
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Nùmero:"
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
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "IDENTIFICACION:"
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
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      DataField       =   "Tipo de identificacion"
      DataSource      =   "Adodc1"
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
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   4800
      Width           =   3015
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
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Apellido:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "FECHA:"
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
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "Nuevo "
         Shortcut        =   ^N
      End
      Begin VB.Menu guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu eliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu REPORTES 
      Caption         =   "REPORTES"
      Begin VB.Menu generar 
         Caption         =   "Generar"
         Shortcut        =   {F7}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Label15.Caption = Time
End Sub

Private Sub Command2_Click()
 Label17.Caption = DateDiff("n", Label12, Label15) & "  minuto(s)"
 
 If Val(Label17) > 0 And Val(Label17) <= 2 Then
 Label19 = "$" & "2000"
 Else
  If Val(Label17) >= 3 And Val(Label17) <= 4 Then
 Label19 = "$" & "4000"
Else
  If Val(Label17) >= 5 And Val(Label17) <= 6 Then
 Label19 = "$" & "6000"
Else
  If Val(Label17) >= 7 And Val(Label17) <= 8 Then
 Label19 = "$" & "8000"
Else
 If Val(Label17) < 1 Then
 MsgBox "No debe nada"
 End If
 End If
 End If
 End If
 End If
 
 
 
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub



Private Sub Command4_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext
End If



End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command7_Click()
Label12.Caption = Time
End Sub

Private Sub eliminar_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Form_Load()



Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo1.AddItem "17"
Combo1.AddItem "18"
Combo1.AddItem "19"
Combo1.AddItem "20"
Combo1.AddItem "21"
Combo1.AddItem "22"
Combo1.AddItem "23"
Combo1.AddItem "24"
Combo1.AddItem "25"
Combo1.AddItem "26"
Combo1.AddItem "27"
Combo1.AddItem "28"
Combo1.AddItem "29"
Combo1.AddItem "30"
Combo1.AddItem "31"


Combo2.AddItem "Enero"
Combo2.AddItem "Febrero"
Combo2.AddItem "Marzo"
Combo2.AddItem "Abril"
Combo2.AddItem "Mayo"
Combo2.AddItem "Junio"
Combo2.AddItem "Julio"
Combo2.AddItem "Agosto"
Combo2.AddItem "Septiembre"
Combo2.AddItem "Octubre"
Combo2.AddItem "Noviembre"
Combo2.AddItem "Diciembre"


Combo3.AddItem "2010"
Combo3.AddItem "2011"
Combo3.AddItem "2012"
Combo3.AddItem "2013"
End Sub

Private Sub generar_Click()
DataReport1.Show
End Sub

Private Sub guardar_Click()
Adodc1.Recordset.Update

End Sub

Private Sub nuevo_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Option6_Click()
Label10 = Option6.Caption

End Sub

Private Sub Option7_Click()
Label10 = Option7.Caption

End Sub

Private Sub salir_Click()
End
End Sub

Private Sub Timer1_Timer()
Label13.Caption = Date
Label7.Caption = Time

End Sub
