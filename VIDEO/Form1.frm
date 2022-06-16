VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   855
   ClientTop       =   615
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   14400
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin MARCHOSOLib.Marchoso M1 
      Height          =   2175
      Left            =   10560
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   3836
      _StockProps     =   1
      FileName        =   ""
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   8280
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AUDIO"
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VIDEO CLIC"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP2 
      Height          =   795
      Left            =   4200
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6350
      _cy             =   1402
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   4215
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   7455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   13150
      _cy             =   7435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VIDEO"
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
wmp1.URL = "E:\Html 2012\video html\metal.wmv"



End Sub

Private Sub Command2_Click()
wmp1.URL = "E:\06 - Polly.mp3"
End Sub

Private Sub Command3_Click()
M1.FileName = "E:\anar.gif"

End Sub

Private Sub Command4_Click()
WMP2.Controls.pause

End Sub



Private Sub Command5_Click()
WMP2.Controls.play

End Sub

Private Sub Form_Load()

WMP2.URL = "E:\blom.mp3"
WMP2.URL = "E:\06 - Polly.mp3"
End Sub

If Combo2 = "The beatles" Then
wmp1.URL = "E:\blom.mp3"
Else
If Combo2 = "Pink Floyd" Then
wmp1.URL = "E:\06 - Polly.mp3"
End If

