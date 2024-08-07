VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
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
 Dim validar As Boolean


If Text1.Text = "" Then
        Label1.Caption = "ingresa fecha bobo"

ElseIf validarCaracter(Text1.Text) = True Then
    dia = separarFecha(Text1, "dia")
    mes = separarFecha(Text1, "mes")
    año = separarFecha(Text1, "año")
    
    If validarFecha(CStr(dia), CStr(mes), CStr(año)) Then
        Label1.Caption = "fecha ingresada" & dia & "/" & mes & "/" & año
    Else
        Label1.Caption = "fecha no valida, tonoto"
    End If


    

    
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
        
    End If
        
End Function

Private Function validarFecha(dia As Integer, mes As Integer, año As Integer) As Boolean



    
    If (mes = 4 Or mes = 6 Or mes = 9 Or mes = 11) And dia >= 1 And dia <= 30 And año >= 1910 Then
        validarFecha = True
    ElseIf (mes = 1 Or mes = 3 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12) And dia >= 1 And dia <= 31 And año >= 1910 Then
        validarFecha = True
    
    Else
        
        validarFecha = False
    
    
    End If
    
    















End Function

