VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   2430
   ClientTop       =   945
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   10185
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1320
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Registradora\bd.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Registradora\bd.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      DataField       =   "N°  Venta"
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
      Left            =   1920
      TabIndex        =   28
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   27
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      DataField       =   "Parte"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      DataField       =   "Tipo de carne"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Caption         =   "CONTAR EN: "
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
      Height          =   2655
      Left            =   5400
      TabIndex        =   5
      Top             =   2160
      Width           =   4575
      Begin VB.TextBox Text4 
         BackColor       =   &H80000006&
         DataField       =   "Kilos"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000006&
         DataField       =   "Libras"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H80000012&
         Caption         =   "KILOS"
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
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000008&
         Caption         =   "LIBRAS"
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
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "DESCRIPCION:"
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
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
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
         Height          =   495
         Left            =   2280
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
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
         Left            =   1320
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9960
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente Venta"
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
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Anterior Venta"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Primer Venta"
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
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ultima Venta"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "CARNICERIA "" LA BUENA ESPERANZA"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   30
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Venta N°"
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
      Height          =   360
      Left            =   480
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      DataField       =   "Total Compra"
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
      Height          =   495
      Left            =   7200
      TabIndex        =   26
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "TOTAL COMPRA:"
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
      TabIndex        =   25
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      DataField       =   "Total Kilos"
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
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "TOTAL KILOS:"
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
      TabIndex        =   23
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Parte:"
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
      Left            =   480
      TabIndex        =   17
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Tipo de      carne:"
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
      Left            =   480
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      DataField       =   "Hora"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "HORA"
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
      Left            =   2880
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "FECHA "
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      DataField       =   "Fecha"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      Caption         =   "TOTAL LIBRAS:"
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
      TabIndex        =   11
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000012&
      DataField       =   "Total libras"
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
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu new 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu dell 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu end 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Reportes"
      Begin VB.Menu gen 
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
 Label12 = Val(Label19) + Val(Label9)

End Sub

Private Sub Command2_Click()
If Val(Text3) > 0 And Val(Text3) <= 1 Then
Label19 = "1000"
Else
If Val(Text3) > 1 And Val(Text3) <= 4 Then
Label19 = "3000"
Else
If Val(Text3) > 4 And Val(Text3) <= 5 Then
Label19 = "4000"
Else
If Val(Text3) > 5 And Val(Text3) <= 7 Then
Label19 = "5500"
Else
If Val(Text3) > 7 And Val(Text3) <= 9 Then
Label19 = "7000"
Else
If Val(Text3) > 9 And Val(Text3) <= 11 Then
Label19 = "9000"
Else
If Val(Text3) > 11 And Val(Text3) <= 12 Then
Label19 = "10000"
Else
If Val(Text3) > 12 And Val(Text3) <= 14 Then
Label19 = "12000"
Else
If Val(Text3) > 14 And Val(Text3) <= 16 Then
Label19 = "14000"
Else
If Val(Text3) > 16 And Val(Text3) <= 20 Then
Label19 = "15000"
Else
If Val(Text3) > 20 And Val(Text3) <= 50 Then
Label19 = "30000"

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
If Val(Text4) > 0 And Val(Text4) <= 1 Then
Label9 = "3000"
Else
If Val(Text4) > 1 And Val(Text4) <= 4 Then
Label9 = "9000"
Else
If Val(Text4) > 4 And Val(Text4) <= 5 Then
Label9 = "12000"
Else
If Val(Text4) > 5 And Val(Text4) <= 7 Then
Label9 = "16500"
Else
If Val(Text4) > 7 And Val(Text4) <= 9 Then
Label9 = "21000"
Else
If Val(Text4) > 9 And Val(Text4) <= 11 Then
Label9 = "27000"
Else
If Val(Text4) > 11 And Val(Text4) <= 12 Then
Label9 = "30000"
Else
If Val(Text4) > 12 And Val(Text4) <= 14 Then
Label9 = "36000"
Else
If Val(Text4) > 14 And Val(Text4) <= 16 Then
Label9 = "42000"
Else
If Val(Text4) > 16 And Val(Text4) <= 20 Then
Label9 = "45000"
Else
If Val(Text4) > 20 And Val(Text4) <= 50 Then
Label9 = "90000"

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If



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

Private Sub dell_Click()
Adodc1.Recordset.Delete

End Sub

Private Sub end_Click()
End
End Sub

Private Sub gen_Click()
DataReport1.Show
End Sub

Private Sub new_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Option6_Click()
If Option6 Then
Label5 = "Libras"
End If

End Sub

Private Sub Option7_Click()
If Option7 Then
Label5 = "Kilos"
End If

End Sub

Private Sub Timer1_Timer()
Label13.Caption = Date
Label7.Caption = Time

End Sub
