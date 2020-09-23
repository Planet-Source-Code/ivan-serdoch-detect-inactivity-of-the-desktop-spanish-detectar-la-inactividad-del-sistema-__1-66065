Attribute VB_Name = "Mouse"
Option Explicit

'Funciones API que se utilizan:
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const VK_LBUTTON = &H1 'Boton izquierdo del Mouse
Public Const VK_RBUTTON = &H2 'Boton derecho del Mouse
Public Const VK_MBUTTON = &H4 'Boton central del Mouse (Generalmente la Ruedita o scroll)

Public Type POINTAPI
    x As Long 'Coordenada X de la posicion del Mouse
    y As Long 'Coordenada Y de la posicion del Mouse
    End Type

Public Function GetX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.x
End Function

Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.y
End Function
