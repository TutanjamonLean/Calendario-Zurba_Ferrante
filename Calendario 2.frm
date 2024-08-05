VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14115
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim dia, mes, año As String
 If validarCaracter(Text1.Text) = True Then
    dia = separarFecha(Text1, "dia")
    mes = separarFecha(Text1, "mes")
    año = separarFecha(Text1, "año")
 
 
 End If
 
End Sub
Private Function validarCaracter(caracter As String) As Boolean

    If Len(caracter) >= 8 Or Len(caracter) <= 10 Then
        validarCaracter = True
    Else
        validarCaracter = False
    End If
    
End Function
Private Function separarFecha(fecha, datofecha As String) As String
Dim dia, mes, año, caract As String
Dim cont, A As Integer

    For A = 1 To Len(fecha)
    
    caract = Mid(fecha, A, 1)
    
    If caract = "/" Then
        cont = cont + 1
    
    ElseIf cont = 0 Then
        dia = dia & caract
    
    ElseIf cont = 1 Then
        mes = mes & caract
    
    ElseIf cont = 2 Then
        año = año & caract
    
    End If
    
    Next A
    
    If datofecha = "dia" Then
        separarFecha = dia
    
    ElseIf datofecha = "mes" Then
        separarFecha = mes
    
    ElseIf datofecha = "año" Then
        separarFecha = año
        
End Function

Private Function validarFecha()
Dim dia, mes, año As Integer

End Function

