VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "TypeText"
   ClientHeight    =   5535
   ClientLeft      =   4380
   ClientTop       =   2070
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7815
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "TypeMachine.frx":0000
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Type"
         Default         =   -1  'True
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   3375
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "TypeMachine.frx":0006
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Del"
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control de Velocidad"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   4935
      Begin ComctlLib.Slider Slider1 
         Height          =   675
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1191
         _Version        =   327682
         Min             =   1
         Max             =   5
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Min"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Velocidad"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Max"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "Click"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   4800
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6480
      Y1              =   5040
      Y2              =   4920
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6480
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   6480
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Muestra texto letra por letra =)"
      Height          =   855
      Left            =   6720
      TabIndex        =   4
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      Height          =   1095
      Index           =   2
      Left            =   6600
      Top             =   4320
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      Height          =   975
      Index           =   1
      Left            =   6120
      Top             =   4320
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      Height          =   975
      Index           =   0
      Left            =   5640
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub AdText(text As String)
Text1.text = Text1.text & text
End Sub
Sub AdLine(text As String)
Text1.text = Text1.text & vbCrLf & text
End Sub

Private Sub Command1_Click()
Text1.text = ""
Text2.text = ""
End Sub


Private Sub Command2_Click()
'PlaySound
Borrar
End Sub

Private Sub Command3_Click()
Dim S As Integer
Select Case Slider1.Value ' Para la velocidad de escritura
Case 1
S = 5
Case 2
S = 4
Case 3
S = 3
Case 4
S = 2
Case 5
S = 1
End Select
If Text2.text = "" Then
TypeMsg "Escriba un mensage primero!", S
Delay (1)
Borrar
Exit Sub
End If
   Text1.text = ""
TypeMsg Text2.text, S
End Sub



Private Sub Form_Activate()
MsgBox App.Path
Text2.SetFocus
Delay (2)
TypeMsg "Muestra un mensage dentro de un TextBox como si se estuviera escribiendo en ese momento :)" & vbCrLf & "By Pinky -> www.frikilabs.tk" & vbCrLf & "For Questions: pinky3009@hotmail.com", 1
End Sub

Private Sub Form_Load()
Text1.MousePointer = 2
Text1.text = ""
Text2.text = ""
Slider1.Value = 5
End Sub
 'Funcion Principal
 
 Sub TypeMsg(Mensage As String, RetrasoLetra As Integer)
'Descomponer mensage en sus X caracteres e imprimir uno por uno cada x segundos
If Len(Mensage) >= 1024 Then Exit Sub ' Si la longitud es mayor de 1024 entonses no hace nada
Dim i As Integer                        ' Que nos da un error de desbordamiento
Dim lac As String
Dim Array_Caracter(1024) As String 'Declaramos un   array de 1024

For i = 1 To Len(Mensage) 'Bucle
    Array_Caracter(i) = Left(Mensage, 1) 'La i variable adquiere la primera letra de la cadena
    Mensage = Right(Mensage, Len(Mensage) - 1) 'Se Borra era letra
    Delay (RetrasoLetra / 20) 'Esperamos unos segundos o milesimas
    If Array_Caracter(i) = " " Then GoTo NoSound 'Si el caracter es un espacion no se reproduce sonido
    PlaySound                                      ' Para darle realismo
NoSound:
    AdText (Array_Caracter(i)) ' Se agrega el i caracter al textbox con una secillita funcion
Next

 End Sub

Private Sub Label3_Click()
Text1.text = ""
Text2.text = ""
Text2.text = "Muestra texto letra por letra =)"
TypeMsg Text2.text, 1
Delay (1)
Text1.text = Left(Text1.text, Len(Text1) - 1)
Delay (1 / 10)
Text1.text = Left(Text1.text, Len(Text1) - 1)
Delay (1 / 2)
TypeMsg ";", 1 / 20
Delay (1 / 20)
TypeMsg ")", 1 / 2
End Sub

Private Sub Slider1_Change() 'Que nos muestre la velocidad
Label1.Caption = ""
Label1.Caption = "Velocidad  " & Slider1.Value
End Sub
Public Sub Delay(ByVal nSecond As Single) ' Debo aclarar, esta funcion no la escribi yo
'Pauses for x seconds.
'Detiene la ejecucion de una rutina o funcion x segundos
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer

      dummy = DoEvents()
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub
 
 
 Sub Borrar()
 Do While (Text1.text <> "")
 Text1.text = Left(Text1.text, Len(Text1) - 1)
 Delay (1 / 50)
 Loop
 End Sub

