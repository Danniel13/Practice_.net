VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form busquedas1 
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10005
   LinkTopic       =   "Form2"
   ScaleHeight     =   6210
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ver Todos Los Registros"
      Height          =   615
      Left            =   7320
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\escolar\bd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "estudiantes"
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "busquedas1.frx":0000
      Height          =   1815
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecciona Bùsqueda Por: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.OptionButton Option3 
         Caption         =   "Turno"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Carrera"
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Matrìcula"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   4
      Top             =   3960
      Width           =   6345
   End
   Begin VB.Menu qeqwe 
      Caption         =   "Archivo"
      Begin VB.Menu trweter 
         Caption         =   "Volver a Opciones"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "busquedas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 2100
MSFlexGrid1.ColWidth(3) = 2500
MSFlexGrid1.ColWidth(4) = 1000
Label2.Visible = False
Text1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
MsgBox "Haz Clic en Archivo", vbInformation, "¡Aviso Importante!"
End Sub
Private Sub volveraopciones_Click()
busquedas1.Hide
Opciones.Show
End Sub

Private Sub Option1_Click()
If Option1 = True Then
Label2.Visible = True
Label2.Caption = "Introduce la Matrícula que buscas"
Text1.Visible = True
Text1 = ""
Text1.SetFocus
End If
End Sub
Private Sub Option2_Click()
If Option2 = True Then
Label2.Visible = True
Label2.Caption = "Introduce la Carrera que buscas"
Text1.Visible = True
Text1 = ""
Text1.SetFocus
End If
End Sub
Private Sub Option3_Click()
If Option3 = True Then
Label2.Visible = True
Label2.Caption = "Introduce el Turno que buscas"
Text1.Visible = True
Text1 = ""
Text1.SetFocus
End If
End Sub
Private Sub Command1_Click()
If Option1 = True Then
Data1.RecordSource = "select * from estudiantes where matricula = " & Val(Text1)
Data1.Refresh
Label1.Visible = True

If Data1.Recordset.EOF Then
MsgBox "La Matrícula: " & Val(Text1) & ", No está en la Base de Datos", vbExclamation, "¡Por Favor Revisa el Número de la Matrícula!"
Text1 = ""
Text1.SetFocus
End If

ElseIf Option2 = True Then
Data1.RecordSource = "select * from estudiantes where carrera = '" & Text1 & "'"
Data1.Refresh
Label1.Visible = True

If Data1.Recordset.EOF Then
MsgBox "La Carrera: '" & Text1 & "'" & " No está en la Base de Datos", vbExclamation, "¡Por Favor Revisa el Nombre de la Carrera!"
Text1 = ""
Text1.SetFocus
End If

ElseIf Option3 = True Then
Data1.RecordSource = "select * from estudiantes where turno = '" & Text1 & "'"
Data1.Refresh
Label1.Visible = True
If Data1.Recordset.EOF Then
MsgBox "El Turno: '" & Text1 & "'" & " No está en la Base de Datos", vbExclamation, "¡Por Favor Revisa el Nombre del Turno!"
Text1 = ""
Text1.SetFocus
End If

End If
Label1 = "Total de Registros de la Consulta: " & (MSFlexGrid1.Rows) - 1 & ""
End Sub

Private Sub Command2_Click()
Text1 = ""
Data1.RecordSource = "estudiantes"
Data1.Refresh
Label1 = "Total de Registros de la búsqueda: " & (MSFlexGrid1.Rows) - 1 & ""
End Sub

Private Sub salir_Click()
End

End Sub

Private Sub trweter_Click()
busqueda.Show

End Sub
