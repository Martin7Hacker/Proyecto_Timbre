VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmprograma 
   BackColor       =   &H00FAFAF5&
   Caption         =   "Virtual Martin Temporize v2017 -  Para Escuelas versi�n Pro."
   ClientHeight    =   7845
   ClientLeft      =   3765
   ClientTop       =   2100
   ClientWidth     =   16695
   Icon            =   "frmprograma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   16695
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Skins"
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      Left            =   13080
      TabIndex        =   50
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   10
      Left            =   13080
      TabIndex        =   49
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   9
      Left            =   13080
      TabIndex        =   48
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   8
      Left            =   13080
      TabIndex        =   47
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   7
      Left            =   13080
      TabIndex        =   46
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   13080
      TabIndex        =   45
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   13080
      TabIndex        =   44
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   13080
      TabIndex        =   43
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   13080
      TabIndex        =   42
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   13080
      TabIndex        =   41
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   13080
      TabIndex        =   40
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdmeses1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   13080
      TabIndex        =   39
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdcod 
      BackColor       =   &H00000000&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      Picture         =   "frmprograma.frx":0CCA
      TabIndex        =   34
      Top             =   2357
      Width           =   255
   End
   Begin VB.CommandButton cmdCalendario 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10870
      TabIndex        =   37
      Top             =   2350
      Width           =   255
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FAFAF5&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   13080
      Pattern         =   "*.skn*"
      TabIndex        =   36
      Top             =   5640
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "frmprograma.frx":1594
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   600
      Top             =   6720
   End
   Begin MSComDlg.CommonDialog cdgAbrir 
      Left            =   4320
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdmasmenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12240
      Picture         =   "frmprograma.frx":17C8
      TabIndex        =   33
      Top             =   200
      Width           =   375
   End
   Begin VB.CommandButton cmdmasmenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   10270
      MaskColor       =   &H00000000&
      Picture         =   "frmprograma.frx":2092
      TabIndex        =   32
      Top             =   200
      Width           =   375
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   6075
      Left            =   9960
      TabIndex        =   30
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picuteMesShop 
      BackColor       =   &H00FBF3E8&
      Height          =   6135
      Left            =   10200
      ScaleHeight     =   6075
      ScaleWidth      =   2715
      TabIndex        =   26
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdSoloDias 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2210
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00DE8F38&
         BorderStyle     =   0  'None
         Height          =   177
         Left            =   0
         ScaleHeight     =   180
         ScaleWidth      =   2415
         TabIndex        =   35
         Top             =   0
         Width           =   2415
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6075
         Left            =   2450
         Max             =   -1
         Min             =   -10
         TabIndex        =   27
         Top             =   0
         Value           =   -1
         Width           =   255
      End
      Begin VB.PictureBox panel1 
         BackColor       =   &H00FBF3E8&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   2535
         TabIndex        =   28
         Top             =   0
         Width           =   2535
         Begin MSComCtl2.MonthView meses 
            Height          =   2460
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   4339
            _Version        =   393216
            ForeColor       =   0
            BackColor       =   16511976
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MonthBackColor  =   16511976
            ShowToday       =   0   'False
            StartOfWeek     =   60358658
            TitleBackColor  =   14585656
            TitleForeColor  =   -2147483639
            TrailingForeColor=   16744576
            CurrentDate     =   41776
         End
      End
   End
   Begin VB.ListBox listiempo 
      Enabled         =   0   'False
      Height          =   645
      Left            =   6240
      TabIndex        =   25
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lisdialogo 
      Enabled         =   0   'False
      Height          =   840
      Left            =   6240
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox liscomando 
      Enabled         =   0   'False
      Height          =   840
      Left            =   6240
      TabIndex        =   23
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   7470
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:39"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "19/09/2017"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            TextSave        =   "N�M"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4058
            MinWidth        =   4058
            Picture         =   "frmprograma.frx":295C
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer autoset 
      Interval        =   10
      Left            =   1080
      Top             =   6720
   End
   Begin VB.ListBox filtro 
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox domingo 
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox sabado 
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox viernes 
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox jueves 
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox miercoles 
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox martes 
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lunes 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   9720
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60358658
      CurrentDate     =   40784
   End
   Begin VB.Frame fram_dias 
      BackColor       =   &H00FAFAF5&
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   10999
      TabIndex        =   4
      ToolTipText     =   "Listado de Progrmaci�n de los dias o el dia que queres activar el timbre."
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Lunes"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Martis"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Miercoles"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Jueves"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Viernes"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "Sabado"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAFAF5&
         Caption         =   "domingo"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1845
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6720
   End
   Begin VB.ListBox listado 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   6000
      Index           =   3
      Left            =   6840
      TabIndex        =   3
      ToolTipText     =   "Pizarr�n de Comentarios ."
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox listado 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   6000
      Index           =   2
      Left            =   4560
      TabIndex        =   2
      ToolTipText     =   "Pizarr�n de Tiempo en segundos ."
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox listado 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   6000
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Pizarr�n de Tipo Entrada o Salida. "
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox listado 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6000
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Pizarr�n de Horarios Programados"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1320
      Picture         =   "frmprograma.frx":4E6E
      Top             =   7680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbllinea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linea de Tiempo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   9360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "&Nuevo..."
         Shortcut        =   ^N
      End
      Begin VB.Menu esp9 
         Caption         =   "-"
      End
      Begin VB.Menu abrir 
         Caption         =   "&Abrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu esp10 
         Caption         =   "-"
      End
      Begin VB.Menu guardard 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu esp11 
         Caption         =   "-"
      End
      Begin VB.Menu guardar 
         Caption         =   "&Guardar como"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu esp12 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu ver 
      Caption         =   "&Ver"
      Begin VB.Menu paneldias 
         Caption         =   "Panel de Dias "
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu reloje 
      Caption         =   ""
   End
   Begin VB.Menu opciones 
      Caption         =   "&Opciones"
      Begin VB.Menu espx0 
         Caption         =   "-"
      End
      Begin VB.Menu registro 
         Caption         =   "Registro"
         Begin VB.Menu nuevot 
            Caption         =   "&Nuevo"
         End
         Begin VB.Menu modificar 
            Caption         =   "&Modificar"
            Shortcut        =   ^M
         End
         Begin VB.Menu eliminacion 
            Caption         =   "&Eliminaci�n"
            Begin VB.Menu eliminartodo 
               Caption         =   "&Eliminar todo..."
               Shortcut        =   ^X
            End
            Begin VB.Menu elimnarseleccionado 
               Caption         =   "&Eliminar seccionado�"
               Shortcut        =   ^E
            End
         End
         Begin VB.Menu desplazar 
            Caption         =   "&Desplazar"
            Begin VB.Menu anterior 
               Caption         =   "<< Anterior"
               Shortcut        =   ^{F8}
            End
            Begin VB.Menu siguiente 
               Caption         =   "Siguiente >>"
               Shortcut        =   ^{F9}
            End
         End
      End
      Begin VB.Menu puerto 
         Caption         =   "&Salida"
         Begin VB.Menu pinesdedatos 
            Caption         =   "&Puerto Paralelo"
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu archivoop 
         Caption         =   "&Opciones de Archivo"
         Begin VB.Menu rutasdearchivo 
            Caption         =   "&Rutas de Archivo"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu automatizarprograma 
         Caption         =   "&Automatizar programa"
         Begin VB.Menu ejecutarcuandoinicieelpc 
            Caption         =   "Ejecutar cuando inicie el PC"
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu usuariodelsoft 
         Caption         =   "&Usuario"
         Begin VB.Menu datospersonales 
            Caption         =   "&Datos personales"
            Shortcut        =   {F12}
         End
      End
      Begin VB.Menu Manualmente 
         Caption         =   "&Utilizar Manualmente"
         Begin VB.Menu timbreliceo 
            Caption         =   "&Dispositivo"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu clendario 
         Caption         =   "&Calendario"
         Shortcut        =   ^I
      End
      Begin VB.Menu generadorderutinas 
         Caption         =   "&Generador de Rutinas de Eventos Programables"
         Shortcut        =   {F2}
      End
      Begin VB.Menu historial 
         Caption         =   "&Historial"
         Visible         =   0   'False
      End
      Begin VB.Menu MoveryPegar 
         Caption         =   "&Mover y Pegar"
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu espx101 
         Caption         =   "-"
      End
   End
   Begin VB.Menu visor 
      Caption         =   "&Visor"
   End
   Begin VB.Menu ventana 
      Caption         =   "&Ventana"
   End
   Begin VB.Menu ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu temasayuda 
         Caption         =   "&Temas de Ayuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu espx 
         Caption         =   "-"
      End
      Begin VB.Menu pIdioma 
         Caption         =   "&Personalizar Idioma"
      End
      Begin VB.Menu acercademicrotime 
         Caption         =   "&Acerca de:"
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu acercadepluins 
         Caption         =   "&Acerca de Virtual Martin temporize"
         Shortcut        =   {F4}
      End
      Begin VB.Menu espx1 
         Caption         =   "-"
      End
      Begin VB.Menu circuitoelectronico 
         Caption         =   "&Circuito Electr�nico"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu definido 
      Caption         =   "definidos"
      Visible         =   0   'False
      Begin VB.Menu mostrar 
         Caption         =   "&Mostrar todos los Meses"
      End
      Begin VB.Menu solodefinidosactuales 
         Caption         =   "&Solo definidos Actuales"
      End
   End
   Begin VB.Menu donativo 
      Caption         =   " ---> &DONATIVO <---"
   End
End
Attribute VB_Name = "frmprograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Principal de Virtual Martin temporize v1.7
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Option Explicit
Private Declare Function SetErrorMode Lib "kernel32" _
(ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203 'DobleClic Izquierdo
Private Const WM_LBUTTONDOWN = &H201 'Clic izquierdo
Private Const WM_RBUTTONUP = &H205 'Clic derecho
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Dim sysTray As NOTIFYICONDATA
Dim Memoria As String
Dim proceso_x As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias _
"LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
ByVal hLibModule As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" _
Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000&
Private m_hMod As Long
Private l_meses As Integer
Private mover(13) As New controles

Private Sub abrir_Click()
 If Me.listado(0).ListCount = 0 Then
 abrirArchivo
 Else
 Select Case MsgBox(lenguaje_Menu(366), vbYesNoCancel + vbInformation)
  Case (vbYes)
  guardarF.Almacenar_Fichero guardar_archivo ' guardo los Datos nuevamente
  borrar.borrar ' destrulle todos los objetos
  sistema.eleminarDatos
  guardar_archivo = ""
  abrirArchivo 'Abre el Archivo nuevamente
  Case (vbNo)
  borrar.borrar ' destrulle todos los objetos
  sistema.eleminarDatos
  guardar_archivo = ""
  abrirArchivo
  Case (vbCancel)
  End Select
 End If
 Unirlistados 0
End Sub

Private Sub abrirArchivo()
 With cdgAbrir
 .DialogTitle = lenguaje_Menu(368) & lenguaje_Menu(376) & ":" & lenguaje_Menu(367)
 .Filter = lenguaje_Menu(368) & lenguaje_Menu(376) & "(*.vmt)|*.vmt|" & lenguaje_Menu(369) & "(*.*)|*.*|"
 .FilterIndex = 1
 .ShowOpen
 If Not (.FileName = "") Then
  If .FileName <> "" Then
   If .CancelError = False Then
   abrirF.Abrir_Fichero .FileName
   guardarF.guardar_archivo = .FileName
   .FileName = ""
 End If
  End If
   End If
 End With
End Sub


Private Sub acercademicrotime_Click()
 frmAcercade.Show 1
End Sub

Private Sub acercadepluins_Click()
 frmAcercade.Show 1
End Sub

Private Sub anterior_Click()
 On Error GoTo no_se
 Dim v As Integer
 v = listado(0).ListIndex
 listado(0).ListIndex = v - 1
 listado(1).ListIndex = v - 1
 listado(2).ListIndex = v - 1
 listado(3).ListIndex = v - 1
no_se:
End Sub

Private Sub Arranque_inicio_pc_Click()
 frmarrancarconwindows.Show 1
End Sub

Private Sub autoset_Timer()
 MonthView1.Value = Date
 devolver_dias
 If listado(0).ListCount = 0 Then
 VScroll2.Visible = False
 ElseIf listado(0).ListCount >= 2 Then
 VScroll2.Visible = True
 End If
End Sub

Private Sub Calendario_Click()
 frmalmanaque.Show 1
End Sub

Private Sub circuitoelectronico_Click()
 frmcircuito.Show 1
End Sub

Private Sub clendario_Click()
 Calendario_Click
End Sub

Private Sub datosdelusuario_Click()
 frmDatos.Show 1
End Sub

Private Sub cmdCalendario_Click()
MostarMD False, True
End Sub

Private Sub cmdcod_Click()
 PopupMenu definido
End Sub

Private Sub cmdmasmenos_Click(Index As Integer)
 Dim dias_m, mes_n As Byte
 Dim mese_s, anio As Integer
 Select Case Index
 Case (0)
 For mese_s = 0 To 11
  mes_n = mes_n + 1
  anio = meses(mese_s).Year + 1
  meses(mese_s).Value = "01/" & "" & mes_n _
  & "" & " / " & "" & anio & ""
 Next mese_s
 despinarTodoslosMeses
 Case (1)
 For mese_s = 0 To 11
  mes_n = mes_n + 1
  anio = meses(mese_s).Year - 1
 If anio = 1752 Then
 MsgBox lenguaje_Menu(370), _
 vbInformation, lenguaje_Menu(368) & lenguaje_Menu(376)
 Exit Sub
 Else
 meses(mese_s).Value = "01/" & "" & _
 mes_n & "" & " / " & "" & anio & ""
 End If
 Next mese_s
 despinarTodoslosMeses
 End Select
End Sub

Private Sub despinarTodoslosMeses()
 Dim dias_x As Byte
 Dim anio_a, anio_c As Integer
 Dim ultimoDiaMes As String
 anio_a = Mid(Date, 7, 10)
 anio_c = meses(0).Year
 For dias_x = 0 To 11
  meses(dias_x).Font.Underline = False
  meses(dias_x).Font.Strikethrough = False
  If anio_a < meses(dias_x).Year Then
  meses(dias_x).Font.Underline = True
  meses(dias_x).Day = 1
  ElseIf anio_a > meses(dias_x).Year Then
  meses(dias_x).Font.Strikethrough = True
  ultimoDiaMes = DateSerial(Year(Now), meses(dias_x).Month + 1, 0)
  ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
  ElseIf anio_a = meses(dias_x).Year Then
  anioIgualaAnio
  End If
 Next dias_x
End Sub

Private Sub mesesDinamicos()
 'tachar dias pasados
 Dim dias As Byte
 Dim ultimoDiaMes As String
 Dim anio As Integer
 For dias = 0 To 11
  meses(dias).Font.Underline = False
  meses(dias).Font.Strikethrough = False
 Next dias
 For dias = 0 To mesDelAnio - 1
  ultimoDiaMes = DateSerial(Year(Now), meses(dias).Month + 1, 0)
  ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
  meses(dias).Day = ultimoDiaMes
  meses(dias).Font.Strikethrough = True
 Next dias
 meses(mesDelAnio).Day = Day(Date)
 For dias = mesDelAnio + 1 To 11
  meses(dias).Day = 1
  meses(dias).Font.Underline = True
 Next dias
End Sub

Private Sub anioIgualaAnio()
 'tachar dias pasados
 Dim dias As Byte
 Dim ultimoDiaMes As String
 Dim anio As Integer
 For dias = 0 To 11
  meses(dias).Font.Underline = False
  meses(dias).Font.Strikethrough = False
  Next dias
 For dias = 0 To mesDelAnio - 1
  ultimoDiaMes = DateSerial(Year(Now), meses(dias).Month + 1, 0)
  ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
  meses(dias).Day = ultimoDiaMes
  meses(dias).Font.Strikethrough = True
 Next dias
 meses(mesDelAnio).Day = Day(Date)
 For dias = mesDelAnio + 1 To 11
  meses(dias).Day = 1
  meses(dias).Font.Underline = True
 Next dias
End Sub

Private Sub cmdmeses1_Click(Index As Integer)
 Dim anio_x As Byte
 With VScroll1
  proceso_x = False
  Select Case Index
  Case 0: .Value = 0
  Case 1: .Value = -1
  Case 2: .Value = -2
  Case 3: .Value = -3
  Case 4: .Value = -4
  Case 5: .Value = -5
  Case 6: .Value = -6
  Case 7: .Value = -7
  Case 8: .Value = -8
  Case 9: .Value = -9
  Case 10
  For anio_x = 0 To 11
  meses(anio_x).Year = Mid(Date, 7, 10) 'meses(1).Year
  Next anio_x
  cmdmeses1_Click mesDelAnio 'se queda en el mes actual
  mesesDinamicos
  End Select
 End With
 
 
 'VScroll2.Left = 70 - listado(3).Left
 
 
End Sub

Private Sub crear_meses()
 Dim meses_d As Byte
 For meses_d = 1 To 12
 l_meses = l_meses + 1
 Load meses(l_meses)
 meses(l_meses).Visible = True
 meses(l_meses).Top = 2280 * l_meses
 panel1.Height = 2280 * l_meses
 With VScroll1
 .Min = 0
 .Max = -l_meses + 3
 End With
 Next
 meses(0).Month = mvwJanuary   'enero
 meses(1).Month = mvwFebruary  'febrero
 meses(2).Month = mvwMarch     'marso
 meses(3).Month = mvwApril     'abril
 meses(4).Month = mvwMay       'mayo
 meses(5).Month = mvwJune      'junio
 meses(6).Month = mvwJuly      'julio
 meses(7).Month = mvwAugust    'agosto
 meses(8).Month = mvwSeptember 'septiembre
 meses(9).Month = mvwOctober   'octubre
 meses(10).Month = mvwNovember 'noviembre
 meses(11).Month = mvwDecember 'diciembre
End Sub







Private Sub cmdSoloDias_Click()
MostarMD True, False
End Sub
Private Sub MostarMD(ByVal dia As Boolean, ByVal mes As Boolean)
fram_dias.Visible = dia
cmdcod.Visible = mes
cmdmasmenos(0).Visible = mes
cmdmasmenos(1).Visible = mes
cmdCalendario.Visible = dia
picuteMesShop.Visible = mes
End Sub



Private Sub datospersonales_Click()
 datosdelusuario_Click
End Sub

Private Sub donativo_Click()
 frmDonativos.Show 1
End Sub

Private Sub ejecutarcuandoinicieelpc_Click()
 Arranque_inicio_pc_Click
End Sub

Private Sub elimartodo_Click()
 Select Case MsgBox(lenguaje_Menu(371), _
 vbYesNo + vbExclamation, lenguaje_Menu(372))
 Case vbYes
 listado(0).Clear
 listado(1).Clear
 listado(2).Clear
 listado(3).Clear
 borrar.borrar ' destrulle los objetos
 sistema.eleminarDatos
 End Select
End Sub

Private Sub eliminartodo_Click()
 elimartodo_Click
End Sub

Private Sub elimnarseleccionado_Click()
 elimniarTimbre_Click
End Sub

Private Sub elimniarTimbre_Click()
 If Not listado(0).ListIndex = -1 Then
  Select Case MsgBox(lenguaje_Menu(373) _
  , vbYesNo + vbInformation, lenguaje_Menu(374))
  Case vbYes
  listado(0).RemoveItem (listado(0).ListIndex)
  listado(1).RemoveItem (listado(1).ListIndex)
  listado(2).RemoveItem (listado(2).ListIndex)
  listado(3).RemoveItem (listado(3).ListIndex)
  lunes(0).RemoveItem (lunes(0).ListIndex)
  martes.RemoveItem (martes.ListIndex)
  miercoles.RemoveItem (miercoles.ListIndex)
  jueves.RemoveItem (jueves.ListIndex)
  viernes.RemoveItem (viernes.ListIndex)
  sabado.RemoveItem (sabado.ListIndex)
  domingo.RemoveItem (domingo.ListIndex)
  Filtro.RemoveItem (Filtro.ListIndex)
  liscomando.RemoveItem (liscomando.ListIndex)
  lisdialogo.RemoveItem (lisdialogo.ListIndex)
  listiempo.RemoveItem (listiempo.ListIndex)
  End Select
  Else
  MsgBox lenguaje_Menu(375) _
  , vbInformation, lenguaje_Menu(368) & lenguaje_Menu(376)
 End If
End Sub

Public Sub File1_Click()
Skin1.LoadSkin App.Path & "\Skins\" & File1.FileName

externosF.guardar_Skins
externosF.Abrir_Skins
'carga Skins con el recurso del formulario requerido
cargar_Skins Me
cmdmeses1(11).Caption = "Skins: " & File1.ListIndex

'Skin1.ApplySkin frmprograma.hwnd
'Skin1.ApplySkin frmAcercade.hwnd
'Skin1.ApplySkin frmalmanaque.hwnd
'Skin1.ApplySkin frmarrancarconwindows.hwnd
'Skin1.ApplySkin frmArranque.hwnd
'Skin1.ApplySkin frmCargarIdioma.hwnd
'Skin1.ApplySkin frmcircuito.hwnd
'Skin1.ApplySkin frmcomentario.hwnd
'Skin1.ApplySkin frmcomo.hwnd
'Skin1.ApplySkin frmDatos.hwnd
'Skin1.ApplySkin frmDonativos.hwnd
'Skin1.ApplySkin frmentradasalida.hwnd
'Skin1.ApplySkin frmfunciones.hwnd
'Skin1.ApplySkin frmhistorial.hwnd
'Skin1.ApplySkin frmidioma.hwnd
'Skin1.ApplySkin frmimpresor.hwnd
'Skin1.ApplySkin frmmensage.hwnd
'Skin1.ApplySkin frmnuevoevento.hwnd
'Skin1.ApplySkin frmopciones.hwnd
'Skin1.ApplySkin frmPrincipioFinal.hwnd
'Skin1.ApplySkin frmprograma.hwnd
'Skin1.ApplySkin frmpuerto.hwnd
'Skin1.ApplySkin frmreloj.hwnd
'Skin1.ApplySkin frmtimbre.hwnd
'Skin1.ApplySkin frmutilizarManual.hwnd
'Skin1.ApplySkin frmVisorEventos.hwnd
End Sub

Private Sub Form_Initialize()
  m_hMod = LoadLibrary("shell32.dll")
End Sub

Private Sub Form_Load()

 Lenguage.definir_lenguage_opciones
 ' Apagar los Puertos de Datos *
 
      puertof.apagar_puertos
 
 '******************************
 frmnuevoevento.devolver_dias
 OcultarP.Ocultar True
 externosF.Abrir_Archivo_Externo
 externosF.Abrir_selecionado
 'registro la estencion del archivo de el programa
 archivoF.CrearAsociacion App.Path & "\" & App.EXEName, _
 "vmt", lenguaje_Menu(368) & lenguaje_Menu(376), App.Path & "\" & "vmt.exe,0" '"util.dll,0"
 abrirExterno
 crear_meses
 frmprograma.WindowState = sistema.ven
 cmdmeses1_Click mesDelAnio
 cmdmeses1_Click 10
 Call cargarIdioma
 ver.Visible = False
 
 File1.Path = "Skins\"
 'File1.ListIndex = 1
 
 externosF.Abrir_Skins
 
 File1_Click
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me
 cmdCalendario.Visible = False
End Sub

Function mesDelAnio()
 mesDelAnio = Mid(Date, 4, 2)
 mesDelAnio = mesDelAnio - 1
End Function

Private Sub abrirExterno()
 abrirF.Abrir_Fichero guardar_archivo
 On Error GoTo no_se
 If guardar_archivo = "" Then
 abrirF.Abrir_Fichero xselecionado
 Else
 If Command$ <> "" Then
 End If
 End If
no_se:
End Sub

Private Sub Form_Resize()
 On Error GoTo no_se
  Dim ubicar As Integer
  For ubicar = 0 To 3
  listado(ubicar).Width = 4000
  listado(ubicar).Height = Me.Height - 1400 '1800
  picuteMesShop.Height = Me.Height - 1800
  lbllinea.Top = listado(0).Top + lbllinea.Top
  'Command1.Top = listado(0).Top
  'Command2.Top = listado(0).Top
  'Command1.Left = listado(0).Left + 500
  'Command2.Left = listado(0).Left
  Next
  VScroll2.Height = listado(0).Height
  VScroll1.Height = listado(0).Height
  picuteMesShop.Height = listado(0).Height
no_se:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 puertof.apagar_puertos
 On Error GoTo no_se:
  If frmprograma.listado(0).ListCount <= 0 Then
  guardar_archivo = ""
  borrar.borrar ' destrullo los objetos
  If guardar_archivo = "" Then
  Quitar_Systray
  End 'cierro todo
  Unload Me
  Quitar_Systray
  End If
  Else
  Cancel = 1 'cancelo cerrar
  frmmensage.Show 1
  End If
no_se:
End Sub

Private Sub generadorderutinas_Click()
 frmentradasalida.Show 1
End Sub

Private Sub guardar_Click()
 With cdgAbrir
 .DialogTitle = lenguaje_Menu(368) & lenguaje_Menu(376) & ":" & lenguaje_Menu(377)
 .Filter = lenguaje_Menu(368) & lenguaje_Menu(376) & "(*.vmt)|*.vmt|" & lenguaje_Menu(369) & "(*.*)|*.*|"
 .FilterIndex = 1
 .FileName = lenguaje_Menu(379)
 .ShowSave
 If .FileName = "" Then
 MsgBox lenguaje_Menu(380), vbInformation
 End If
 If .FileName <> "" Then
 If .CancelError = False Then
 guardarF.Almacenar_Fichero .FileName
 guardar_archivo = .FileName
 Else
 End If
 End If
 End With
End Sub

Public Sub guardard_Click()
 If guardar_archivo = "" Then
 guardar_Click
 Else
 guardarF.Almacenar_Fichero guardar_archivo
 End If
End Sub

Private Sub historial_Click()
frmhistorial.Show 1
End Sub



Private Sub listado_Click(Index As Integer)
 Unirlistados Index
 On Error GoTo no_se
 seleccionarLista listado(0).ListIndex
 seleccionarLista listado(1).ListIndex
 seleccionarLista listado(2).ListIndex
 seleccionarLista listado(3).ListIndex
 VScroll2.Value = listado(0).ListIndex
 VScroll2.Value = listado(1).ListIndex
 VScroll2.Value = listado(2).ListIndex
 VScroll2.Value = listado(3).ListIndex
no_se:
End Sub

Private Sub listado_DragDrop(Index As Integer, Source _
 As Control, X As Single, Y As Single)
 Unirlistados Index
End Sub

Private Sub modificar_Click()
 On Error GoTo no_se
 Dim c As Integer
 For c = 0 To 300
 selecionar_dias
 With frmnuevoevento
 .boton(1).Caption = lenguaje_Menu(225)
 MemoriaF.dias = True
 .labinfo.Caption = lenguaje_Menu(226)
 .DTPicker1.Value = listado(0).List(listado(0).ListIndex)
 .Combo1(0).Text = listado(1).List(listado(1).ListIndex)
 .Combo1(1).Text = listado(2).List(listado(2).ListIndex)
 .Combo1(2).ListIndex = Filtro.List(Filtro.ListIndex)
 .Text1.Text = listado(3).List(listado(3).ListIndex)
 End With
 Next
 frmnuevoevento.Show 1
no_se:
End Sub

Private Sub selecionar_dias()
 Const lunesx As String = "2"     'lunes
 Const martesx As String = "3"    'martes
 Const miercolesx As String = "4" 'miercoles
 Const juevesx As String = "5"    'jueves
 Const viernesx As String = "6"   'viernes
 Const sabadox As String = "7"    'sabados
 Const domingox As String = "1"   'domingos
 With frmnuevoevento
 'lunes
 If lunes(0).List(lunes(0).ListIndex) = lunesx Then
 .Check1(0).Value = lunes(0).List(lunes(0).ListIndex) - 1
 End If
 'martes
 If martes.List(martes.ListIndex) = martesx Then
 .Check1(1).Value = martes.List(lunes(0).ListIndex) - 2
 End If
 'miercoles
 If miercoles.List(miercoles.ListIndex) = miercolesx Then
 .Check1(2).Value = miercoles.List(miercoles.ListIndex) - 3
 End If
 'jueves
 If jueves.List(jueves.ListIndex) = juevesx Then
 .Check1(3).Value = jueves.List(jueves.ListIndex) - 4
 End If
 'viernes
 If viernes.List(viernes.ListIndex) = viernesx Then
 .Check1(4).Value = viernes.List(viernes.ListIndex) - 5
 End If
 'sabados
 If sabado.List(sabado.ListIndex) = sabadox Then
 .Check1(5).Value = sabado.List(sabado.ListIndex) - 6
 End If
 'domnigo
 If domingo.List(domingo.ListIndex) = domingox Then
 .Check1(6).Value = domingo.List(domingo.ListIndex)
 End If
 End With
End Sub

Private Sub modificart_Click()
 modificar_Click
End Sub

Private Sub mostrar_Click()
 proceso_x = True
End Sub

Private Sub MoveryPegar_Click()
 moverControles
End Sub

Private Sub nuevo_Click()
 With frmnuevoevento
 MemoriaF.dias = False
 .Show 1
 .labinfo.Caption = Lenguage.lenguaje_Menu(207)
 .boton(1).Caption = Lenguage.lenguaje_Menu(224)
 End With
End Sub

Private Sub nuevot_Click()
 nuevo_Click
End Sub

Private Sub obsgen_Click()
 frmpuerto.Show 1
End Sub

Private Sub paneldias_Click()
 If paneldias.Checked = False Then
 paneldias.Checked = True
 fram_dias.Visible = True
 ElseIf paneldias.Checked = True Then
 paneldias.Checked = False
 fram_dias.Visible = False
 End If
End Sub

Private Sub pIdioma_Click()
frmidioma.Show 1
End Sub

Private Sub pinesdedatos_Click()
 frmpuerto.Show 1
End Sub

Private Sub reloje_Click()
 On Error GoTo no_se
 frmreloj.Show 1
no_se:
End Sub

Private Sub rutas_Click()
 frmArranque.Show 1
 
End Sub

Private Sub rutasdearchivo_Click()
 rutas_Click
End Sub

Private Sub salir_Click()
 Form_Unload True
End Sub

Private Sub siguiente_Click()
 On Error GoTo no_se
 Dim v As Integer
 v = listado(0).ListIndex
 listado(0).ListIndex = v + 1
 listado(1).ListIndex = v + 1
 listado(2).ListIndex = v + 1
 listado(3).ListIndex = v + 1
no_se:
End Sub

Private Sub solodefinidosactuales_Click()
 proceso_x = False
End Sub

Private Sub temasayuda_Click()
'  MsgBox "Por haora no existe archivo de Ayuda", _
'  vbInformation, "Archivos de Ayuda"
frmcomentario.Show 1
End Sub

Private Sub timbreliceo_Click()
 utilizarmaual_Click
End Sub

Private Sub Timer1_Timer()
 reloje.Caption = Time
End Sub

Private Sub Unirlistados(ByVal union As Integer)
On Error GoTo no_se
 Dim uni As Integer
 For uni = 0 To 3
 listado(uni).ListIndex = listado(union).ListIndex
 Next uni
no_se:
End Sub

Private Sub listado_DblClick(Index As Integer)
 Unirlistados Index
End Sub

Private Sub listado_DragOver(Index As Integer, _
 Source As Control, X As Single, Y As Single, State As Integer)
 Unirlistados Index
End Sub

Private Sub listado_GotFocus(Index As Integer)
 Unirlistados Index
End Sub

Private Sub listado_ItemCheck(Index As Integer, _
 Item As Integer)
 Unirlistados Index
End Sub

Private Sub listado_KeyDown(Index As Integer, _
 KeyCode As Integer, Shift As Integer)
 Unirlistados Index
End Sub

Private Sub listado_KeyPress(Index As Integer, _
 KeyAscii As Integer)
 Unirlistados Index
End Sub

Private Sub listado_KeyUp(Index As Integer, KeyCode _
 As Integer, Shift As Integer)
 Unirlistados Index
End Sub

Private Sub listado_LostFocus(Index As Integer)
 Unirlistados Index
End Sub

Private Sub listado_MouseDown(Index As Integer, Button _
 As Integer, Shift As Integer, X As Single, Y As Single)
 Unirlistados Index
 If Button = vbRightButton Then
 PopupMenu opciones ' muestra un men� deslizable en pantalla
 End If
End Sub

Private Sub listado_MouseMove(Index As Integer, Button _
 As Integer, Shift As Integer, X As Single, Y As Single)
 Unirlistados Index
End Sub

Private Sub listado_MouseUp(Index As Integer, Button _
 As Integer, Shift As Integer, X As Single, Y As Single)
 Unirlistados Index
End Sub

Private Sub listado_OLECompleteDrag(Index As Integer, Effect _
 As Long)
 Unirlistados Index
End Sub

Private Sub listado_OLEDragDrop(Index As Integer, Data As DataObject _
 , Effect As Long, Button As Integer, Shift As Integer, X As Single, _
 Y As Single)
 Unirlistados Index
End Sub

Private Sub listado_OLEDragOver(Index As Integer, Data As DataObject _
 , Effect As Long, Button As Integer, Shift As Integer, X As Single, Y _
 As Single, State As Integer)
 Unirlistados Index
End Sub

Private Sub listado_OLEGiveFeedback(Index As Integer, Effect _
 As Long, DefaultCursors As Boolean)
 Unirlistados Index
End Sub

Private Sub listado_OLESetData(Index As Integer, Data As DataObject _
 , DataFormat As Integer)
 Unirlistados Index
End Sub

Private Sub listado_OLEStartDrag(Index As Integer, Data As DataObject _
 , AllowedEffects As Long)
 Unirlistados Index
End Sub

Private Sub listado_Scroll(Index As Integer)
 Unirlistados Index
 seleccionarLista listado(0).ListIndex
End Sub

Private Sub listado_Validate(Index As Integer _
 , Cancel As Boolean)
 Unirlistados Index
End Sub

Private Sub si_abro_archivo()
 If Not (externosF.xselecionado = "") Then
 abrirF.Abrir_Fichero externosF.xselecionado
 guardar_archivo = externosF.xselecionado
 End If
End Sub

Private Sub registrar()
 On Error GoTo no_se
 If Command$ <> "" Then
 End If
no_se:
End Sub

Private Sub utilizarmaual_Click()
frmutilizarManual.Show 1
End Sub

Private Sub seleccionarLista(ByVal indice As Integer)
 lunes(0).ListIndex = indice
 martes.ListIndex = indice
 miercoles.ListIndex = indice
 jueves.ListIndex = indice
 viernes.ListIndex = indice
 sabado.ListIndex = indice
 domingo.ListIndex = indice
 Filtro.ListIndex = indice
 liscomando.ListIndex = indice
 lisdialogo.ListIndex = indice
 listiempo.ListIndex = indice
End Sub
Public Sub devolver_dias()
 'lunes
 Select Case lunes(0).List(lunes(0).ListIndex)
  Case (2)
  Check1(0).Value = 1
  Case (0)
  Check1(0).Value = 0
 End Select
 'martes
 Select Case martes.List(martes.ListIndex)
  Case (3)
  Check1(1).Value = 1
  Case (0)
  Check1(1).Value = 0
 End Select
 'miercoles
 Select Case miercoles.List(miercoles.ListIndex)
  Case (4)
  Check1(2).Value = 1
  Case (0)
  Check1(2).Value = 0
 End Select
 'jueves
 Select Case jueves.List(jueves.ListIndex)
  Case (5)
  Check1(3).Value = 1
  Case (0)
  Check1(3).Value = 0
 End Select
 'viernes
 Select Case viernes.List(viernes.ListIndex)
  Case (6)
  Check1(4).Value = 1
  Case (0)
  Check1(4).Value = 0
 End Select
 'sabado
 Select Case sabado.List(sabado.ListIndex)
  Case (7)
  Check1(5).Value = 1
  Case (0)
  Check1(5).Value = 0
 End Select
 'domingo
 Select Case domingo.List(domingo.ListIndex)
  Case (1)
  Check1(6).Value = 1
  Case (0)
  Check1(6).Value = 0
 End Select
End Sub

Private Sub colocar_icono_en_la_bandeja(intervalo As Integer)
 'Datos varios de la estructura
 With sysTray
 .cbSize = Len(sysTray)
 ' -- Establecer el Hwnd de la ventana
 .hwnd = Me.hwnd
 ' -- Definir el handle de la barra de tarea (identificador)
 .uId = 1&
 ' -- Establecer los flags para la estructura
 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 ' -- Establecer el mensaje Callback a Windows
 .uCallBackMessage = WM_LBUTTONDOWN
 ' -- Establecer el Picture a hicon
 .hIcon = Me.Icon
 ' -- Establecer el ToolTip
 .szTip = lenguaje_Menu(368) & lenguaje_Menu(376) & Chr$(0)
 End With
 ' -- llamar a Shell_NotifyIcon para Crear y agregar el icono
 Call Shell_NotifyIcon(NIM_ADD, sysTray)
 ' -- Ocultar el Formulario
 Me.Hide
 ' -- Inicializar el temporizador
 Timer1.Interval = intervalo
End Sub

Private Sub Quitar_Systray()
 With sysTray
 .cbSize = Len(sysTray)
 .hwnd = Me.hwnd
 .uId = 1&
 End With
 ' -- Le pasamos el mensaje NIM_DELETE para eliminar el programa del �rea de notificaci�n
 Call Shell_NotifyIcon(NIM_DELETE, sysTray)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift _
 As Integer, X As Single, Y As Single)
 Dim Msg
 Msg = X / Screen.TwipsPerPixelX
 ' -- Si hacemos DobleClick Izquierdo ..
 If Msg = WM_LBUTTONDBLCLK Then
 Me.Show
 ' -- Desplegar el PopUp menu
 ElseIf Msg = WM_RBUTTONUP Then
 Me.PopupMenu archivo
 End If
End Sub

Public Sub mostrar_menu(ByVal mostrar As Boolean)
 With opciones
 .Enabled = mostrar
 .Visible = mostrar
 End With
 With archivo
 .Enabled = mostrar
 .Visible = mostrar
 End With
 With ver
 .Enabled = mostrar
 .Visible = mostrar
 End With
 With reloje
 .Enabled = mostrar
 .Visible = mostrar
 End With
 With opciones
 .Enabled = mostrar
 .Visible = mostrar
 End With
 With ayuda
 .Enabled = mostrar
 .Visible = mostrar
 End With
End Sub

Private Sub Timer2_Timer()
 disparar.disparar
End Sub

Private Sub ventana_Click()
 frmcomo.Show 1
End Sub

Private Sub visor_Click()
 frmVisorEventos.Show 1
End Sub

Private Sub VScroll1_Change()
 panel1.Top = VScroll1.Value * 2280
 If proceso_x = True And VScroll1.Value = 0 Then
 VScroll1.Value = -8
 cmdmasmenos_Click 0
 despinarTodoslosMeses
 End If
 If proceso_x = True And VScroll1.Value = -9 Then
 VScroll1.Value = -1
 cmdmasmenos_Click 1
 despinarTodoslosMeses
 End If
End Sub

Private Sub VScroll1_Scroll()
 VScroll1_Change
End Sub

Private Sub VScroll2_Change()
 On Error GoTo nose
 With VScroll2
 .Max = listado(0).ListCount - 1
 .Min = 0
 listado(0).ListIndex = .Value
 listado(1).ListIndex = .Value
 listado(2).ListIndex = .Value
 listado(3).ListIndex = .Value
 End With
nose:
End Sub

Private Sub VScroll2_Scroll()
 VScroll2_Change
End Sub

Private Sub moverControles()
 mover(0).moverDato listado(0)
 mover(1).moverDato listado(1)
 mover(2).moverDato listado(2)
 mover(3).moverDato listado(3)
End Sub

Private Sub moverOtros()
 mover(3).moverDato lunes(0)
 mover(4).moverDato martes
 mover(5).moverDato miercoles
 mover(6).moverDato jueves
 mover(7).moverDato viernes
 mover(8).moverDato sabado
 mover(9).moverDato domingo
 mover(10).moverDato Filtro
 mover(11).moverDato lisdialogo
 mover(12).moverDato listiempo
 mover(13).moverDato liscomando
End Sub

Public Sub cargarIdioma()
With frmprograma
.archivo.Caption = lenguaje_Menu(0)
.nuevo.Caption = lenguaje_Menu(1)
.abrir.Caption = lenguaje_Menu(2)
.guardard.Caption = lenguaje_Menu(3)
.guardar.Caption = lenguaje_Menu(4)
.salir.Caption = lenguaje_Menu(5)
.ver.Caption = lenguaje_Menu(6)
.paneldias.Caption = lenguaje_Menu(7)
.opciones.Caption = lenguaje_Menu(8)
.registro.Caption = lenguaje_Menu(9)
.nuevot.Caption = lenguaje_Menu(10)
.modificar.Caption = lenguaje_Menu(11)
.eliminacion.Caption = lenguaje_Menu(12)
.eliminartodo.Caption = lenguaje_Menu(13)
.elimnarseleccionado.Caption = lenguaje_Menu(14)
.desplazar.Caption = lenguaje_Menu(15)
.anterior.Caption = lenguaje_Menu(16)
.siguiente.Caption = lenguaje_Menu(17)
.puerto.Caption = lenguaje_Menu(18)
.pinesdedatos.Caption = lenguaje_Menu(19)
.archivoop.Caption = lenguaje_Menu(20)
.rutasdearchivo.Caption = lenguaje_Menu(21)
.automatizarprograma.Caption = lenguaje_Menu(22)
.ejecutarcuandoinicieelpc.Caption = lenguaje_Menu(23)
.usuariodelsoft.Caption = lenguaje_Menu(24)
.datospersonales.Caption = lenguaje_Menu(25)
.Manualmente.Caption = lenguaje_Menu(26)
.timbreliceo.Caption = lenguaje_Menu(27)
.clendario.Caption = lenguaje_Menu(28)
.generadorderutinas.Caption = lenguaje_Menu(29)
.visor.Caption = lenguaje_Menu(30)
.ventana.Caption = lenguaje_Menu(31)
.ayuda.Caption = lenguaje_Menu(32)
.temasayuda.Caption = lenguaje_Menu(33)
.pIdioma.Caption = lenguaje_Menu(34)
.acercademicrotime.Caption = lenguaje_Menu(35)
.acercadepluins.Caption = lenguaje_Menu(36)
.circuitoelectronico.Caption = lenguaje_Menu(37)
.donativo.Caption = lenguaje_Menu(38)
.mostrar.Caption = lenguaje_Menu(39)
.solodefinidosactuales.Caption = lenguaje_Menu(40)
.historial.Caption = lenguaje_Menu(41)

listado(0).ToolTipText = lenguaje_Menu(241)
listado(1).ToolTipText = lenguaje_Menu(242)
listado(2).ToolTipText = lenguaje_Menu(243)
listado(3).ToolTipText = lenguaje_Menu(244)

cmdmeses1(0).Caption = lenguaje_Menu(245)
cmdmeses1(1).Caption = lenguaje_Menu(246)
cmdmeses1(2).Caption = lenguaje_Menu(247)
cmdmeses1(3).Caption = lenguaje_Menu(248)
cmdmeses1(4).Caption = lenguaje_Menu(249)
cmdmeses1(5).Caption = lenguaje_Menu(250)
cmdmeses1(6).Caption = lenguaje_Menu(251)
cmdmeses1(7).Caption = lenguaje_Menu(252)
cmdmeses1(8).Caption = lenguaje_Menu(253)
cmdmeses1(9).Caption = lenguaje_Menu(254)
cmdmeses1(10).Caption = lenguaje_Menu(255)
 
End With
End Sub
