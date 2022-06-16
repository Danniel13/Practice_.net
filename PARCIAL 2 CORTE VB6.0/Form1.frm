VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   510
   ClientTop       =   225
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   15870
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Imagen del disco:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   600
      TabIndex        =   22
      Top             =   3480
      Width           =   4935
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000012&
         Caption         =   "# de disco"
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
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000012&
         Caption         =   "Artista:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "FACTURA: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6975
      Left            =   6120
      TabIndex        =   9
      Top             =   2160
      Width           =   6015
      Begin VB.ListBox List3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   360
         ItemData        =   "Form1.frx":0000
         Left            =   1920
         List            =   "Form1.frx":0002
         TabIndex        =   30
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   360
         ItemData        =   "Form1.frx":0004
         Left            =   1920
         List            =   "Form1.frx":0006
         TabIndex        =   29
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   360
         ItemData        =   "Form1.frx":0008
         Left            =   1920
         List            =   "Form1.frx":000A
         TabIndex        =   28
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Facturar"
         Height          =   375
         Left            =   4320
         TabIndex        =   26
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1695
         Left            =   3120
         TabIndex        =   16
         Top             =   4200
         Width           =   2655
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000012&
            Caption         =   "Otro"
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
            Left            =   480
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton a 
            BackColor       =   &H80000012&
            Caption         =   "Crédito"
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
            Left            =   480
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton b 
            BackColor       =   &H80000012&
            Caption         =   "Efectivo"
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
            Left            =   480
            TabIndex        =   17
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   11
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   10
         Top             =   1830
         Width           =   2295
      End
      Begin VB.Label Label14 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   31
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000012&
         Caption         =   "Artista"
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
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         Caption         =   "Forma de Pago: "
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
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Disco #:"
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
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Fecha: "
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
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Height          =   480
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   480
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000012&
         Caption         =   "Ciudad:"
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
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   7680
      TabIndex        =   7
      Text            =   "Dìa"
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   8760
      TabIndex        =   6
      Text            =   "Mes"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   10440
      TabIndex        =   5
      Text            =   "Año"
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   465
      Left            =   2400
      TabIndex        =   3
      Text            =   "Disco "
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   465
      Left            =   2400
      TabIndex        =   2
      Text            =   "Artista"
      Top             =   1560
      Width           =   2535
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
      Left            =   6480
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "# de disco  "
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
      Height          =   480
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Artista:  "
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
      Height          =   480
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "DISCOTIENDA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
Label14 = a.Caption

End Sub

Private Sub b_Click()
Label14 = b.Caption

End Sub

Private Sub Combo2_Change()
If Combo2.Text = "The beatles" Then
Text7 = "The beatles"
End If

End Sub

Private Sub Command1_Click()

List3.AddItem Combo6 & " de " & Combo5 & " de " & Combo4

List2.AddItem Combo3
List1.AddItem Combo2

If Combo2 = "The beatles" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\beat.jpg")
Else
If Combo2 = "Pink Floyd" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\pink.jpg")
Else
If Combo2 = "Nirvana" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\nir.jpg")
Else
If Combo2 = "Hector Lavoe" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\lavoe.jpg")
Else
If Combo2 = "Wilfredo Vargas" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\wil.jpg")
Else
If Combo2 = "Jorge Velosa" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\jorg.jpg")
Else
If Combo2 = "Guns & Roses" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\guns.jpg")
Else
If Combo2 = "Syd Barret" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\sid.jpg")
Else
If Combo2 = "Deff Leppard" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\def.jpg")
Else
If Combo2 = "Binomio de Oro" Then
Image1.Picture = LoadPicture("F:\PARCIAL 2 CORTE\bin.jpg")
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

If Combo3 = "1" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\1.jpg")
Else
If Combo3 = "2" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\2.jpg")
Else
If Combo3 = "3" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\3.jpg")
Else
If Combo3 = "4" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\4.jpg")
Else
If Combo3 = "5" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\5.jpg")
Else
If Combo3 = "6" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\6.jpg")
Else
If Combo3 = "7" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\7.jpg")
Else
If Combo3 = "8" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\8.jpg")
Else
If Combo3 = "9" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\9.jpg")
Else
If Combo3 = "10" Then
Image2.Picture = LoadPicture("F:\PARCIAL 2 CORTE\10.jpg")
Else

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

Private Sub Form_Load()
Combo2.AddItem "The beatles"
Combo2.AddItem "Pink Floyd"
Combo2.AddItem "Nirvana"
Combo2.AddItem "Hector Lavoe"
Combo2.AddItem "Wilfredo Vargas"
Combo2.AddItem "Jorge Velosa"
Combo2.AddItem "Guns & Roses"
Combo2.AddItem "Syd Barret"
Combo2.AddItem "Deff Leppard"
Combo2.AddItem "Binomio de Oro"



Combo3.AddItem "1"
Combo3.AddItem "2"
Combo3.AddItem "3"
Combo3.AddItem "4"
Combo3.AddItem "5"
Combo3.AddItem "6"
Combo3.AddItem "7"
Combo3.AddItem "8"
Combo3.AddItem "9"
Combo3.AddItem "10"



Combo6.AddItem "1"
Combo6.AddItem "2"
Combo6.AddItem "3"
Combo6.AddItem "4"
Combo6.AddItem "5"
Combo6.AddItem "6"
Combo6.AddItem "7"
Combo6.AddItem "8"
Combo6.AddItem "9"
Combo6.AddItem "10"
Combo6.AddItem "11"
Combo6.AddItem "12"
Combo6.AddItem "13"
Combo6.AddItem "14"
Combo6.AddItem "15"
Combo6.AddItem "16"
Combo6.AddItem "17"
Combo6.AddItem "18"
Combo6.AddItem "19"
Combo6.AddItem "20"
Combo6.AddItem "21"
Combo6.AddItem "22"
Combo6.AddItem "23"
Combo6.AddItem "24"
Combo6.AddItem "25"
Combo6.AddItem "26"
Combo6.AddItem "27"
Combo6.AddItem "28"
Combo6.AddItem "29"
Combo6.AddItem "30"
Combo6.AddItem "31"


Combo5.AddItem "Enero"
Combo5.AddItem "Febrero"
Combo5.AddItem "Marzo"
Combo5.AddItem "Abril"
Combo5.AddItem "Mayo"
Combo5.AddItem "Junio"
Combo5.AddItem "Julio"
Combo5.AddItem "Agosto"
Combo5.AddItem "Septiembre"
Combo5.AddItem "Octubre"
Combo5.AddItem "Noviembre"
Combo5.AddItem "Diciembre"

Combo4.AddItem "2000"
Combo4.AddItem "2001"
Combo4.AddItem "2002"
Combo4.AddItem "2003"
Combo4.AddItem "2004"
Combo4.AddItem "2005"
Combo4.AddItem "2006"
Combo4.AddItem "2007"
Combo4.AddItem "2008"
Combo4.AddItem "2009"
Combo4.AddItem "2010"
Combo4.AddItem "2011"
Combo4.AddItem "2012"
Combo4.AddItem "2013"
End Sub

Private Sub Option1_Click()
If Option1 Then
MsgBox "No se acepta otra forma de pago", 16, "Atención:"
End If

End Sub
