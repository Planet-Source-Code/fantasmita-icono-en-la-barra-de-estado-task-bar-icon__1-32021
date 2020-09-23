'Creado por: El Fantasmita & El Chacal
'Lugar: En algún lugar del mundo
'Fecha: Jueves, 21 de Febrero del 2002
'  WEB: http://www.geocities.com/utileria

VERSION 5.00
Begin VB.Form Formulario 
   Caption         =   "Formulario para ver un icono en la barra de tareas"
   ClientHeight    =   6195
   ClientLeft      =   1875
   ClientTop       =   1920
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6195
   ScaleWidth      =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "&Minimizar / Minized"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton BotSal 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3000
      TabIndex        =   0
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Menu mPopUpSys 
      Caption         =   "&Systray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BotSal_Click()
    vgComando = "¿Realmente Deseas Salir?" & Chr(13) & Chr(13) & _
        "Si minimizas este formulario podrás ver un icono en la barra de tareas... junto al reloj"
    If vbYes = MsgBox(vgComando, vbYesNo, "¡Aviso!") Then
        End
    End If
End Sub

Private Sub Command1_Click()
  Me.WindowState = 1
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As _
         Single, Y As Single)
      'Este procedimiento recibe las llamadas del SySTray (System Tray Icon)en la barra de tareas.
      Dim result As Long
      Dim msg As Long
      'El valor de X variará dependiendo de la configuración
      msg = x / Screen.TwipsPerPixelX
      
      Select Case msg
        Case WM_LBUTTONDBLCLK    '515 Restaura la ventana de Windows

          'Si al restaurar este formulario quieres que este formulario aparezca maximizado debes quitar la opcion vbNormal y debes poner vbMaximized
          Me.WindowState = vbMaximized  '<-- Prueba también con vbMinimized ó vbNormal
          result = SetForegroundWindow(Me.hwnd)
          Me.Show
    
          With nid
              .cbSize = Len(nid)
              .hwnd = Me.hwnd
              .uId = vbNull
              .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
              .uCallBackMessage = WM_MOUSEMOVE
              .hIcon = Me.Icon
              .szTip = "Titulo Formulario" & vbNullChar
          End With
          Shell_NotifyIcon NIM_DELETE, nid
    
        Case WM_RBUTTONUP        '517 Desplegando Menú Popup
         result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopUpSys
       End Select
End Sub

Private Sub Form_Resize()
    Dim vlPos As Long
    Dim vltam As Long
    
    If Me.WindowState <> 1 Then
        If Me.Width < 7140 Then
           Me.Width = 6810
        End If
        If Me.Height < 2191 Then
           Me.Height = 2191
        End If
        
        vltam = 0
        vltam = vltam + (180 * 6) ' Mas Espacios Intermedios
        vlPos = (Me.Width - vltam) / 2
        
    Else
     With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Estas mirando el icono en la barra de tareas de mi formulario" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    
    Me.Hide
    End If
End Sub

Private Sub mPopExit_Click()
    Unload Me
    End
End Sub

Private Sub mPopRestore_Click()
    Dim result As Long

    'Si al restaurar este formulario quieres que este formulario aparezca maximizado debes quitar la opcion vbNormal y debes poner vbMaximized
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub


