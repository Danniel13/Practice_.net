VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "INSERTA EL NUMERO Y APRENDE COMO SE ESCRIBE"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "PROCURA QUE NO PASE DE 9 CIFRAS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "EL NUMERO EN LETRAS:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "INTRODUCE EL NUMERO, PRESIONA ENTER,  Y DATE CUENTA COMO SE ESCRIBE:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim val_Dec As String, X As String, a As String, n As Integer
Dim posic As Integer, ante As String, post As String
Dim masculino As Boolean

Function NumberToString(ByVal number As Double, ByVal masculino As Boolean) As String
'


    If number = 0 Then
        NumberToString = "Cero"
        Exit Function
    End If
    If number < 0 Then
        number = number * -1
    End If
    X = CStr(Fix(number))          ' ...entero,
    Do While Len(X) < 9         ' ...y de 9 cifras.
        X = "0" & X
    Loop
    a = ""
' Grupos de 3 cifras, de atrás hacia adelante:
    For n = 7 To 1 Step -3
' ¿El grupo actual es cero?:
        If CInt(Mid(X, n, 3)) <> 0 Then
' No.
' Tratar casos especiales decena:
            Select Case CInt(Mid(X, n + 1, 2))
                Case 10
                    a = "diez " & a
                Case 11
                    a = "once " & a
                Case 11
                    a = "once " & a
                Case 12
                    a = "doce " & a
                Case 13
                    a = "trece " & a
                Case 14
                    a = "catorce " & a
                Case 15
                    a = "quince " & a
                Case 16
                    a = "dieciseis " & a
                Case 17
                    a = "diecisiete " & a
                Case 18
                    a = "dieciocho " & a
                Case 19
                    a = "diecinueve " & a
                Case 20
                    a = "veinte " & a
                Case 21
                    If n > 1 Then
                        a = "veintiun$ " & a
                    Else
                        a = "veintiun " & a
                    End If
                Case 22
                    a = "veintidos " & a
                Case 23
                    a = "veintitrés " & a
                Case 24
                    a = "veinticuatro " & a
                Case 25
                    a = "veinticinco " & a
                Case 26
                    a = "veintiseis " & a
                Case 27
                    a = "veintisiete " & a
                Case 28
                    a = "veintiocho " & a
                Case 29
                    a = "veintinueve " & a
                Case Else
' Restantes casos; traducir unidad:
                    Select Case CInt(Mid(X, n + 2, 1))
                        Case 0
                        Case 1
                            Select Case n
                                Case 7
                                    a = "y un$ " & a
                                Case 4
                                    If masculino Then
                                        a = "y un " & a
                                    Else
                                        a = "y una " & a
                                    End If
                                Case 1
                                    a = "y un " & a
                            End Select
                        Case 2
                            a = "y dos " & a
                        Case 3
                            a = "y tres " & a
                        Case 4
                            a = "y cuatro " & a
                        Case 5
                            a = "y cinco " & a
                        Case 6
                            a = "y seis " & a
                        Case 7
                            a = "y siete " & a
                        Case 8
                            a = "y ocho " & a
                        Case 9
                            a = "y nueve " & a
                    End Select
' Traducir decena:
                    Select Case CInt(Mid(X, n + 1, 1))
                        Case 0
                        Case 3
                            a = "treinta " & a
                        Case 4
                            a = "cuarenta " & a
                        Case 5
                            a = "cincuenta " & a
                        Case 6
                            a = "sesenta " & a
                        Case 7
                            a = "setenta " & a
                        Case 8
                            a = "ochenta " & a
                        Case 9
                            a = "noventa " & a
                    End Select
            End Select
' Prever caso "ciento y tres":
            If Left(a, 1) = "y" Then
                a = Right(a, Len(a) - 2)
            End If
' Traducir centena:
            Select Case CInt(Mid(X, n, 1))
                Case 0
                Case 1
                    If CInt(Mid(X, n + 1, 2)) = 0 Then
                        a = "cien " & a
                    Else
                        a = "ciento " & a
                    End If
                Case 2
                    a = "doscient$s " & a
                Case 3
                    a = "trescient$s " & a
                Case 4
                    a = "cuatrocient$s " & a
                Case 5
                    a = "quinient$s " & a
                Case 6
                    a = "seiscient$s " & a
                Case 7
                    a = "setecient$s " & a
                Case 8
                    a = "ochocient$s " & a
                Case 9
                    a = "novecient$s " & a
                
            End Select
        End If
' Poner terminación del grupo anterior:
' Puede haber quedado "y tres":
        If Left(a, 1) = "y" Then
            a = Right(a, Len(a) - 2)
        End If
' Millones:
        If n = 4 Then
            If CInt(Left(X, 3)) = 1 Then
                a = "millón " & a
            Else
                If CInt(Left(X, 3)) <> 0 Then
                    a = "millones " & a
                End If
            End If
        Else
            If n = 7 Then
' Miles:
                If CInt(Mid(X, 4, 3)) = 1 Then
                    a = "mil " & a
                Else
                    If CInt(Mid(X, 4, 3)) <> 0 Then
                        a = "mil " & a
                    End If
                End If
            End If
        End If
' Traducir género, "$" en el texto. Para el grupo de los
' millones, se traduce siempre por masculino:
        If n = 1 Then
            masculino = True
        End If
        posic = 1
        Do While posic <> 0
            posic = InStr(a, "$")
            If posic <> 0 Then
                ante = Left(a, posic - 1)
                post = Right(a, Len(a) - posic)
                If masculino Then
                    a = ante & "o" & post
                Else
                    a = ante & "a" & post
                End If
            End If
        Loop
    Next n
' Caso especial: puede haber quedado "unx mil "
    If Left(a, 7) = "un mil " Then
        a = Right(a, Len(a) - 3)
    Else
        If Left(a, 8) = "una mil " Then
            a = Right(a, Len(a) - 4)
        End If
    End If
' Inicial en mayúsculas:
    If val_Dec = "" Then
    NumberToString = Trim(UCase(Left(a, 1)) & Right(a, Len(a) - 1))
    Else
    NumberToString = Trim(Left(a, 1) & Right(a, Len(a) - 1))
    End If
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)


If IsNumeric(Text1) Then
If KeyAscii = 13 Then

    Text2 = NumberToString(Text1, True)

    If InStr(Text1, ",") > 0 Then
        val_Dec = Mid(Text1, InStr(Text1, ",") + 1, Len(Text1))
        Text2 = Text2 & " con " & NumberToString(val_Dec, True) & " céntimos"
    End If

    
End If
End If

End Sub

