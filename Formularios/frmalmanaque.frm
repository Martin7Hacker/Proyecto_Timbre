VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmalmanaque 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Martin temporize: Calendario"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "frmalmanaque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechaHoy 
      Caption         =   "&Ir a la fecha de Hoy"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -1200
      Picture         =   "frmalmanaque.frx":0CCA
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   1
      Top             =   0
      Width           =   8160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Calendario Grafico"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FBF3E8&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   480
      ScaleHeight     =   525
      ScaleWidth      =   1620
      TabIndex        =   0
      Top             =   4680
      Width           =   1620
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   4710
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   8308
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16511976
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
Attribute VB_Name = "frmalmanaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Calendario Grafico para el programa Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Private Sub cmdsalir_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdFechaHoy_Click()
 MonthView1(0).Value = Date
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmalmanaque
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cmdFechaHoy_Click
 Call cargarIdioma
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub




Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(317)
Label1.Caption = lenguaje_Menu(318)
cmdFechaHoy.Caption = lenguaje_Menu(319)
cmdSalir.Caption = lenguaje_Menu(320)
End Sub

Private Sub MonthView1_KeyPress(Index As Integer, KeyAscii As Integer)
salir_op KeyAscii
End Sub
