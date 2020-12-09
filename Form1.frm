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
   Begin VB.Label DIVISIÓN 
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
Private Sub CERO_Click()                    'Se trabajará en el botón "cero"
RESULTADO = RESULTADO + "0"                 'Esta función mostrará en el objeto RESULTADO el número 0 varías veces
End Sub                                     'Finaliza nuestro trabajo en el botón "cero"
Private Sub CINCO_Click()
RESULTADO = RESULTADO + "5"                 'Se guarda el número 5 en el objeto RESULTADO y si se vuelve a seleccionar ese botón aparecerá otro 5 en RESULTADO
End Sub
Private Sub CUATRO_Click()
RESULTADO = RESULTADO + "4"                 'Esta función mostrará en el objeto RESULTADO el número 4 varías veces
End Sub
Private Sub DIVISIÓN_Click()                'Se trabajará en el botón división
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la función división
RESULTADO = Empty
TIPO_OPER = 2                               'El valor designado para la operación división será el número 2
End Sub
Private Sub DOS_Click()
RESULTADO = RESULTADO + "2"                 'Esta función mostrará en el objeto RESULTADO el número 2 varías veces
End Sub
Private Sub IGUAL_Click()                   'Se trabajará con el botón igual
If TIPO_OPER = 1 Then                       'Si el tipo de operación seleccionada se coloca en 1 se hará la siguiente línea
RESULTADO = VAR1 * Val(RESULTADO)           'Nuestra var1 se multiplicará con el valor registrado en val(Resultado) y se mostrará en Resultado
End If

If TIPO_OPER = 2 Then                       'La operación seleccionada se llevará a cabo
RESULTADO = VAR1 / Val(RESULTADO)           'Nuestra var1 se dividirá con el valor registrado en val(Resultado) y se mostrará en Resultado
End If

If TIPO_OPER = 3 Then                       'La operación designada con el valor 3 se realizará en la siguiente línea
RESULTADO = VAR1 + Val(RESULTADO)           'Nuestra var1 se sumará con el valor registrado en val(Resultado) y se mostrará en Resultado
End If

If TIPO_OPER = 4 Then                       'La condición seleccionada con el tipo de operación 4 se llevará a cabo
RESULTADO = VAR1 - Val(RESULTADO)           'Nuestra var1 se restará con el valor registrado en val(Resultado) y se mostrará en Resultado
End If

If TIPO_OPER = 5 Then
RESULTADO = VAR1 - ((VAR1 * Val(RESULTADO)) / 100)   'Nuestra var1 se multiplicará con el valor registrado en val(Resultado) y se dividirá entre 100 para luego ser mostrada en Resultado
End If

End Sub
Private Sub MULTIPLICAR_Click()             'Se trabajará en el objeto multiplicar
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la función MULTIPLICAR
RESULTADO = Empty
TIPO_OPER = 1                               'La operación multiplicar estará designada por el valor 1
End Sub
Private Sub NUEVE_Click()
RESULTADO = RESULTADO + "9"                 'Esta función mostrará en el objeto RESULTADO el número 9 varías veces
End Sub
Private Sub OCHO_Click()
RESULTADO = RESULTADO + "8"                 'Esta función mostrará en el objeto RESULTADO el número 8 varías veces
End Sub
Private Sub PORCENTAJE_Click()              'Se trabajará en el objeto Porcentaje
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la función PORCENTAJE
RESULTADO = Empty
TIPO_OPER = 5                               'El valor asignado para esta operación es el número 5
End Sub
Private Sub RESTA_Click()                   'Se trabajará en el objeto Resta
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la función RESTA
RESULTADO = Empty
TIPO_OPER = 4                               'Nuestra operación Resta estará designada con el valor 4
End Sub
Private Sub SEIS_Click()
RESULTADO = RESULTADO + "6"                 'Esta función mostrará en el objeto RESULTADO el número 6 varías veces
End Sub
Private Sub SIETE_Click()
RESULTADO = RESULTADO + "7"                 'Esta función mostrará en el objeto RESULTADO el número 7 varías veces
End Sub
Private Sub SUMA_Click()                    'Se trabajara en el objeto SUMA
VAR1 = Val(RESULTADO)                       'Se guarda en VAR1 el valor que se marque en RESULTADO para despues limpiar la pantalla y realizar la función SUMA
RESULTADO = Empty
TIPO_OPER = 3                               'Se declara el número de operación, esta es nuestra operación #3, se enuncia con numeros para evitar confusiones en un futuro
End Sub                                     'Se termina de trabajar en este objeto
Private Sub TRES_Click()
RESULTADO = RESULTADO + "3"                 'Esta función mostrará en el objeto RESULTADO el número 3 varías veces
End Sub
Private Sub UNO_Click()
RESULTADO = RESULTADO + "1"                 'Esta función mostrará en el objeto RESULTADO el número 1 varías veces
End Sub
