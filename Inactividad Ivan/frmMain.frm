VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   1710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContador 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Seg de Inactividad"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer tmrClicks 
      Interval        =   1
      Left            =   840
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NuevaPosX As Integer
Private NuevaPosY As Integer

Private ViejaPosX As Integer
Private ViejaPosY As Integer

Private Contador As Byte

Private Sub Form_Load()

'Poner las cordenadas iniciales del mouse
ViejaPosX = GetX
ViejaPosY = GetY

'Establecer el contador a cero
Contador = 0
End Sub

Private Sub Timer1_Timer()

'Establecer la nueva posicion
NuevaPosX = GetX
NuevaPosY = GetY

'Conpararla con la vieja
    If ViejaPosX = NuevaPosX And ViejaPosY = NuevaPosY Then
        Contador = Contador + 1
    Else
        Contador = 0
        ViejaPosX = NuevaPosX
        ViejaPosY = NuevaPosY
    End If

txtContador.Text = Contador 'Esto solo es para referencia, se puede borrar...

'Si el contador llega al tope...
    If Contador = 10 Then
        MsgBox "Pasaron 10 segundos de inactividad del mouse!"
        'Aca va la acción!!!
        Contador = 0
    End If
End Sub

Private Sub tmrClicks_Timer()
'Si se hace click con el mouse (click derecho, izquierdo y central) vuelve el contador a cero
    If GetAsyncKeyState(VK_LBUTTON) Or GetAsyncKeyState(VK_RBUTTON) Or GetAsyncKeyState(VK_MBUTTON) Then
        Contador = 0
        Exit Sub 'Sale porque ya detecto actividad (en los botones del mouse); no es necesario fijarse si se presiono una tecla del teclado
    End If

'Si no se hizo click, se fija si se presiona una tecla del teclado
    Dim i As Integer
    For i = 0 To 255
        If GetAsyncKeyState(i) = -32767 Then
            Contador = 0
            Exit For
        End If
    Next
End Sub
