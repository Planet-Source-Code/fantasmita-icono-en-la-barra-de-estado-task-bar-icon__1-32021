'Creado por: El Fantasmita & El Chacal
'Lugar: En algún lugar del mundo
'Fecha: Jueves, 21 de Febrero del 2002
'  WEB: http://www.geocities.com/utileria

Attribute VB_Name = "Icono"
Option Explicit

Public vgComando As String      ' Comandos en General

'El usuario define el requerimiento de la llamada del Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Constantes requiridas por la llamada del API Shell_NotifyIcon:
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDBLCLK = &H203   'Doble clic
Public Const WM_RBUTTONUP = &H205       'Botón Arriba (Desplega el Menú)

Public nid As NOTIFYICONDATA

#If Win32 Then
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
#End If

