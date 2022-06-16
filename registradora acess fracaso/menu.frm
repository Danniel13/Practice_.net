VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H80000007&
   Caption         =   "MENU"
   ClientHeight    =   630
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MENU 
      Caption         =   "MENU"
      Begin VB.Menu RES 
         Caption         =   "RES"
      End
      Begin VB.Menu POLLO 
         Caption         =   "POLLO"
      End
      Begin VB.Menu CERDO 
         Caption         =   "CERDO"
      End
      Begin VB.Menu PESCADO 
         Caption         =   "PESCADO"
      End
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CERDO_Click()
form3.Show

End Sub

Private Sub PESCADO_Click()
form4.Show


End Sub

Private Sub POLLO_Click()
form2.Show


End Sub

Private Sub RES_Click()
form1.Show


End Sub
