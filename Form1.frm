VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3885
   FillColor       =   &H008080FF&
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox RESULTADO 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "@Adobe Ming Std L"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label PORCENTAJE 
      Caption         =   " % "
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label IGUAL 
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label DIVISI�N 
      Caption         =   " /"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   14
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label MULTIPLICAR 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   13
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label RESTA 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label SUMA 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label CERO 
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   10
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label NUEVE 
      Caption         =   " 9"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label OCHO 
      Caption         =   " 8"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   8
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label SIETE 
      AutoSize        =   -1  'True
      Caption         =   " 7"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label SEIS 
      Caption         =   " 6"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label CINCO 
      Caption         =   " 5"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label CUATRO 
      Caption         =   " 4"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label TRES 
      Caption         =   " 3"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label DOS 
      Caption         =   " 2"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label UNO 
      Caption         =   " 1"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VAR1 As Integer                      'se declaran las variables
Public TIPO_OPER As Integer
Private Sub CERO_Click()                    'Se trabajar� en el bot�n "cero"
RESULTADO = RESULTADO + "0"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 0 var�as veces
End Sub                                     'Finaliza nuestro trabajo en el bot�n "cero"
Private Sub CINCO_Click()
RESULTADO = RESULTADO + "5"                 'Se guarda el n�mero 5 en el objeto RESULTADO y si se vuelve a seleccionar ese bot�n aparecer� otro 5 en RESULTADO
End Sub
Private Sub CUATRO_Click()
RESULTADO = RESULTADO + "4"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 4 var�as veces
End Sub
Private Sub DIVISI�N_Click()                'Se trabajar� en el bot�n divisi�n
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la funci�n divisi�n
RESULTADO = Empty
TIPO_OPER = 2                               'El valor designado para la operaci�n divisi�n ser� el n�mero 2
End Sub
Private Sub DOS_Click()
RESULTADO = RESULTADO + "2"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 2 var�as veces
End Sub
Private Sub IGUAL_Click()                   'Se trabajar� con el bot�n igual
If TIPO_OPER = 1 Then                       'Si el tipo de operaci�n seleccionada se coloca en 1 se har� la siguiente l�nea
RESULTADO = VAR1 * Val(RESULTADO)           'Nuestra var1 se multiplicar� con el valor registrado en val(Resultado) y se mostrar� en Resultado
End If

If TIPO_OPER = 2 Then                       'La operaci�n seleccionada se llevar� a cabo
RESULTADO = VAR1 / Val(RESULTADO)           'Nuestra var1 se dividir� con el valor registrado en val(Resultado) y se mostrar� en Resultado
End If

If TIPO_OPER = 3 Then                       'La operaci�n designada con el valor 3 se realizar� en la siguiente l�nea
RESULTADO = VAR1 + Val(RESULTADO)           'Nuestra var1 se sumar� con el valor registrado en val(Resultado) y se mostrar� en Resultado
End If

If TIPO_OPER = 4 Then                       'La condici�n seleccionada con el tipo de operaci�n 4 se llevar� a cabo
RESULTADO = VAR1 - Val(RESULTADO)           'Nuestra var1 se restar� con el valor registrado en val(Resultado) y se mostrar� en Resultado
End If

If TIPO_OPER = 5 Then
RESULTADO = VAR1 - ((VAR1 * Val(RESULTADO)) / 100)   'Nuestra var1 se multiplicar� con el valor registrado en val(Resultado) y se dividir� entre 100 para luego ser mostrada en Resultado
End If

End Sub
Private Sub MULTIPLICAR_Click()             'Se trabajar� en el objeto multiplicar
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la funci�n MULTIPLICAR
RESULTADO = Empty
TIPO_OPER = 1                               'La operaci�n multiplicar estar� designada por el valor 1
End Sub
Private Sub NUEVE_Click()
RESULTADO = RESULTADO + "9"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 9 var�as veces
End Sub
Private Sub OCHO_Click()
RESULTADO = RESULTADO + "8"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 8 var�as veces
End Sub
Private Sub PORCENTAJE_Click()              'Se trabajar� en el objeto Porcentaje
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la funci�n PORCENTAJE
RESULTADO = Empty
TIPO_OPER = 5                               'El valor asignado para esta operaci�n es el n�mero 5
End Sub
Private Sub RESTA_Click()                   'Se trabajar� en el objeto Resta
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la funci�n RESTA
RESULTADO = Empty
TIPO_OPER = 4                               'Nuestra operaci�n Resta estar� designada con el valor 4
End Sub
Private Sub SEIS_Click()
RESULTADO = RESULTADO + "6"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 6 var�as veces
End Sub
Private Sub SIETE_Click()
RESULTADO = RESULTADO + "7"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 7 var�as veces
End Sub
Private Sub SUMA_Click()                    'Se trabajara en el objeto SUMA
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la funci�n SUMA
RESULTADO = Empty
TIPO_OPER = 3                               'Se declara el n�mero de operaci�n, esta es nuestra operaci�n #3, se enuncia con numeros para evitar confusiones en un futuro
End Sub                                     'Se termina de trabajar en este objeto
Private Sub TRES_Click()
RESULTADO = RESULTADO + "3"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 3 var�as veces
End Sub
Private Sub UNO_Click()
RESULTADO = RESULTADO + "1"                 'Esta funci�n mostrar� en el objeto RESULTADO el n�mero 1 var�as veces
End Sub
