VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Alumnos 
   BackColor       =   &H80000007&
   Caption         =   "Sistema Escolar"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   15240
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "F:\escolar\bd.mdb"
      DefaultCursorType=   1  'ODBCCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "estudiantes"
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      DataField       =   "turno"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "carrera"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "matrìcula"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1935
      Left            =   3000
      TabIndex        =   0
      Top             =   4920
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Carrera"
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Turno"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Matrìcula"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Menu Movimientos 
      Caption         =   "Movimientos"
      Begin VB.Menu Nuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Buscar 
         Caption         =   "Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu Alumnos 
         Caption         =   "Alumnos"
      End
   End
End
Attribute VB_Name = "Alumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alumnos_Click()
DataReport1.Show
End Sub

Private Sub Buscar_Click()
busqueda.Show
End Sub

Private Sub Eliminar_Click()
If MsgBox("¿Quieres Eliminar la Matrícula Número: " & Text1 & "?", 16 + 4) = 6 Then
Data1.Recordset.Delete
Data1.Refresh
Text1.SetFocus
MsgBox "Se Eliminó la Matrícula", vbCritical, "Aviso Importante"
Else
MsgBox "No se Eliminó la Matrícula Número: " & Text1, vbExclamation, "Aviso Importante"
End If

End Sub

Private Sub Form_activate()
With MSFlexGrid1
For X = 1 To .Rows - 1
.Row = X
For J = 1 To .Cols - 1
.Col = J
.CellBackColor = IIf((X Mod 2) = 1, Val(&HC0FFFF), Val(&HC0FFC0))
.CellFontBold = True
.CellForeColor = &HFF0000
Next J
Next X
End With
End Sub

Private Sub Form_Load()
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 1100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
MsgBox "Haz Clic en Movimientos", vbInformation, "¡Aviso Importante!"
End Sub

Private Sub Guardar_Click()
Data1.UpdateRecord
Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
End Sub

Private Sub Nuevo_Click()
Data1.Recordset.AddNew
End Sub

Private Sub salir_Click()
End

End Sub
