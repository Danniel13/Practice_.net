VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   12075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox m 
      Height          =   855
      Left            =   7440
      TabIndex        =   9
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MAYOR"
      Height          =   975
      Left            =   2760
      TabIndex        =   8
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox c 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox b 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox a 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EL MAYOR DE TRES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   405
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   2880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EL MAYOR ES: "
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
      Height          =   405
      Left            =   1800
      TabIndex        =   3
      Top             =   6120
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESE EL TERCER NUMERO"
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
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Top             =   5040
      Width           =   4365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESE EL SEGUNDO NUMERO"
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
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   3960
      Width           =   4710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESE EL PRIMER NUMERO"
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
      Height          =   405
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   4365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(a) > Val(b) And Val(a) > Val(c) Then
m = a
Else

If Val(b) > Val(c) And Val(b) > Val(a) Then
m = b
Else


If Val(c) > Val(b) And Val(c) > Val(a) Then
m = c



End If

End If


End If


End Sub

