VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMemoTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MemoTest"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmMemoTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrDelay 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   3600
      Top             =   4200
   End
   Begin VB.Timer TmrTiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   4200
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   4575
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   28
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   27
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   26
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   25
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   24
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   23
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   22
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   21
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   20
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   19
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   18
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   17
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   16
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   15
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   14
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   13
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   12
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   11
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   10
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   9
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   8
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   7
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   6
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   5
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   4
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   3
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkMemo 
      ForeColor       =   &H8000000F&
      Height          =   975
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblTiempo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tiempo: 0 Segundos."
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   29
      Top             =   4200
      Width           =   2325
   End
   Begin VB.Menu MnuJuego 
      Caption         =   "&Juego"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Linea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu MnuSonido 
      Caption         =   "&Sonido"
      Begin VB.Menu MnuEfectos 
         Caption         =   "&Efectos"
         Checked         =   -1  'True
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MnuAcerca 
      Caption         =   "&Acerca"
   End
End
Attribute VB_Name = "FrmMemoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de Variables y Vectores
Dim A, I, X, Z, R, rVec(28) As Integer, Ficha(28) As String
Dim E(2), C, P, G, T, D As Integer

'Función Para Reproducción de Sonidos
Private Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10

Private Sub ChkMemo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MnuEfectos.Checked = True Then SndPlaySound App.Path & "\Sonidos\GirderImpact.Wav", _
    SND_ASYNC Or SND_NODEFAULT 'Reproducir Sonido
    
    If T = 0 Then TmrTiempo.Enabled = True 'Habilitar Timer de Conteo Solo La 1ra Vez
    If ChkMemo(Index).Value = 1 Then
        C = C + 1 'Contar Fichas Clickeadas
        If C = 1 Then 'Si Se Clickea La Primera Ficha
            E(1) = Index
        ElseIf C = 2 Then 'Si Se Clickea La Segunda Ficha
            E(2) = Index
            C = 0 'Resetear Contador de Clicks
            D = 1 'Valor Del Delay
            TmrDelay.Enabled = True 'Comenzar Delay
            For I = 1 To 28
                ChkMemo(I).Enabled = False
            Next I
            ChkMemo(E(1)).Enabled = True
            ChkMemo(E(2)).Enabled = True
        End If
    Else
        ChkMemo(Index).Value = 1
    End If
End Sub

Private Sub Comparacion()
    If Val(ChkMemo(E(1)).Tag) = Val(ChkMemo(E(2)).Tag) Then 'Si Las Imágenes Son Iguales
        If MnuEfectos.Checked = True Then SndPlaySound App.Path & "\Sonidos\Laser.Wav", _
        SND_ASYNC Or SND_NODEFAULT 'Reproducir Sonido
        ChkMemo(E(1)).Visible = False
        ChkMemo(E(2)).Visible = False
        G = G + 1 'Sumar Pares Ganados
    Else 'Si Las Imágenes Son Distintas
        If MnuEfectos.Checked = True Then SndPlaySound App.Path & "\Sonidos\MineArm.Wav", _
        SND_ASYNC Or SND_NODEFAULT 'Reproducir Sonido
        ChkMemo(E(1)).Value = 0 'Poner Cara Abajo Las 2 Fichas Clickeadas
        ChkMemo(E(2)).Value = 0 '  "
        P = P + 1 'Sumar Pares Perdidos
    End If
    
    StatusBar.SimpleText = "Jugadas Realizadas: " & G & " Par(es) Exitoso(s); " & P & " Intento(s) Fallido(s)."
    If G = 14 Then 'Si Se Completó El Tablero
        TmrTiempo.Enabled = False 'Deshabilitar Timer de Conteo
        
        'Mensaje de Felicitación
        MsgBox "¡Felicitaciones! Ha Ganado Esta Partida En " & G + P & " Jugadas y " & T & " Segundos." & _
        vbCrLf & vbCrLf & "     Presione F2 Para Iniciar Un Nuevo Juego.", vbInformation, "Partida Ganada!"
    End If
    
    For I = 1 To 28
        ChkMemo(I).Enabled = True
    Next I
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Graficos\F1.Jpg") 'Cargar Fondo
    StatusBar.SimpleText = "¡Bienvenido a MemoTest! - Presione F2 Para Iniciar Un Nuevo Juego."
    
    For I = 1 To 28
        'Cargar Dorso de las Fichas
        ChkMemo(I).Picture = LoadPicture(App.Path & "\Graficos\D1.Jpg")
        ChkMemo(I).DisabledPicture = LoadPicture(App.Path & "\Graficos\D1.Jpg")
    Next I
End Sub

Private Sub MnuAcerca_Click()
    FrmAcerca.Show
End Sub

Private Sub MnuEfectos_Click()
    If MnuEfectos.Checked = True Then
        MnuEfectos.Checked = False
    Else
        MnuEfectos.Checked = True
    End If
End Sub

Private Sub MnuMusica_Click()
    If MnuMusica.Checked = True Then
        MnuMusica.Checked = False
    Else
        MnuMusica.Checked = True
    End If
End Sub

Private Sub MnuNuevo_Click()
    If MnuEfectos.Checked = True Then SndPlaySound App.Path & "\Sonidos\ElectricTransp.Wav", _
    SND_ASYNC Or SND_NODEFAULT 'Reproducir Sonido
    StatusBar.SimpleText = "Se Ha Iniciado Un Nuevo Juego."
    
    TmrTiempo.Enabled = False 'Deshabilitar Timer de Conteo
    LblTiempo.Caption = "Tiempo: 0 Segundos." 'Actualizar Label De Tiempo
    C = 0 'Resetear Contador de Fichas Clickeadas
    G = 0 'Resetear Contador de Pares Ganados
    P = 0 'Resetear Contador de Pares Perdidos
    X = 1 'Evitar Que Se Tome El Elemento 0 del Contador
    T = 0 'Resetear El Contador De Segundos
            
    For I = 1 To 28 'Resetear rVec() Para Incorporar Nuevos Randoms
        rVec(I) = 0
    Next I
    
    'Determinar Número de CheckBox al Azar en rVec() Sin Repetición
    Do Until X = 29 '29 Contrarresta El Efecto De X = 1
        A = 0 'Setear Marcador en 0
        Randomize
        R = Int(Rnd * 28 + 1)
        For Z = 1 To 28
            If rVec(Z) = R Then A = 1
        Next Z
        If A <> 1 Then
            rVec(X) = R
            X = X + 1
        End If
    Loop
    
    'Guardar Imágenes de las Fichas en el Vector y Cargarlas en los CheckBox
    For I = 1 To 28
        ChkMemo(I).Value = 0 'Dar Vuelta Todas Las Fichas
        ChkMemo(I).Visible = True 'Mostrar Todas Las Fichas
        If I <= 14 Then
            Ficha(I) = App.Path & "\Graficos\" & I & ".Jpg"
            ChkMemo(rVec(I)).DownPicture = LoadPicture(Ficha(I))
            ChkMemo(rVec(I)).Tag = I 'Almacenar En Cada Tag El Número de Imagen
        ElseIf I >= 15 And I <= 28 Then
            Ficha(I) = App.Path & "\Graficos\" & I - 14 & ".Jpg"
            ChkMemo(rVec(I)).DownPicture = LoadPicture(Ficha(I))
            ChkMemo(rVec(I)).Tag = I - 14 'Almacenar En Cada Tag El Número de Imagen
        End If
    Next I
End Sub

Private Sub MnuSalir_Click()
    End
End Sub

Private Sub TmrDelay_Timer()
    D = D - 1
    If D = 0 Then
        Call Comparacion 'Ejecutar Comparación De Fichas Después Del Delay
        TmrDelay.Enabled = False
    End If
End Sub

Private Sub TmrTiempo_Timer()
    T = T + 1
    LblTiempo.Caption = "Tiempo: " & T & " Segundos."
End Sub
