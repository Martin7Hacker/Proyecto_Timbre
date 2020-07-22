VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmfunciones 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "funciones al sistema"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmfunciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   120
      Width           =   6255
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Funciones de Sistemas Operables."
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
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4005
      End
   End
   Begin VB.CommandButton cmdcomentarios 
      Caption         =   "comentarios:"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox cob1 
      BackColor       =   &H00FBF3E8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmfunciones.frx":0CCA
      Left            =   4320
      List            =   "frmfunciones.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   3495
         TabIndex        =   8
         Top             =   240
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   2295
            Left            =   0
            Picture         =   "frmfunciones.frx":0CCE
            Top             =   0
            Width           =   3525
         End
      End
      Begin VB.Frame frame2 
         BackColor       =   &H00EDAC85&
         Height          =   2055
         Left            =   4200
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   420
            Left            =   2040
            TabIndex        =   4
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   0
            CustomFormat    =   "m"
            Format          =   119472131
            UpDown          =   -1  'True
            CurrentDate     =   0.805555555555556
         End
         Begin VB.TextBox txtd 
            BackColor       =   &H00EDAC85&
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
            Height          =   1815
            Left            =   120
            MaxLength       =   127
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label labinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Sin dialogo..."
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1680
            TabIndex        =   6
            Top             =   840
            Width           =   1785
         End
         Begin VB.Label lbld 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo ="
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
            Height          =   195
            Left            =   1200
            TabIndex        =   5
            Top             =   960
            Width           =   795
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2535
         Left            =   3960
         ScaleHeight     =   2475
         ScaleWidth      =   0
         TabIndex        =   1
         Top             =   120
         Width           =   60
      End
   End
   Begin HookMenu.XpMenu XpMenu12 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   17
      BmpCount        =   1
      CheckBorderColor=   7021576
      SelMenuBorder   =   0
      SelMenuBackColor=   16511976
      SelMenuForeColor=   0
      SelCheckBackColor=   12367532
      MenuBorderColor =   0
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   0
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   14215660
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   4194368
      ArrowNormalColor=   14585656
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "frmfunciones.frx":1B434
      Key:1           =   "#mc:0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu comentarios 
      Caption         =   "comentarios"
      Visible         =   0   'False
      Begin VB.Menu mc 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmfunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Funciones para Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Public devolver_comando As String

Private Sub cargar_controles()
 With cob1
 .AddItem Lenguage.lenguaje_Menu(174)
 .AddItem Lenguage.lenguaje_Menu(175)
 .AddItem Lenguage.lenguaje_Menu(176)
 .AddItem Lenguage.lenguaje_Menu(177)
 .AddItem Lenguage.lenguaje_Menu(178)
 .AddItem Lenguage.lenguaje_Menu(179)
 .AddItem Lenguage.lenguaje_Menu(180)
 .AddItem Lenguage.lenguaje_Menu(181)
 End With
End Sub

Private Sub cmdAplicar_Click()
devolverString
sistema.tomarDatos
 frmnuevoevento.Text1.Text = txtd.Text & lenguaje_Menu(275) & DTPicker1.Minute
 frmnuevoevento.Combo1(1).Text = DTPicker1.Minute
 If cmdAplicar.Caption = Lenguage.lenguaje_Menu(225) Then
 sistema.modificarDatos 'modifica los datos ingresado
 'frmprograma.liscomando.List(frmprograma.liscomando.ListIndex) = devolver_comando
 
 End If
 If cob1.Text = "" Then
 MsgBox lenguaje_Menu(276)
 End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdcomentarios_Click()
 PopupMenu comentarios
End Sub

Private Sub cmdcomentarios_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsel_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cob1_Click()
 If cob1.ListIndex = 5 Then
 labinfo.Visible = False
 Frame2.Visible = True
 txtd.Visible = False
 lbld.Visible = True
 DTPicker1.Visible = True
 ElseIf cob1.ListIndex = 6 Then
 labinfo.Visible = False
 Frame2.Visible = True
 txtd.Visible = True
 DTPicker1.Visible = False
 lbld.Visible = False
 Else
 labinfo.Visible = True
 Frame2.Visible = True
 txtd.Visible = False
 DTPicker1.Visible = False
 lbld.Visible = False
 End If
End Sub

Private Sub devolverString()
 On Error GoTo no_se
 Select Case cob1.ListIndex
  Case (0)
  devolver_comando = ""   '// sin opcion
  Case (1)
  devolver_comando = "so.dll -s -f" '// Apagar el equipo
  Case (2)
  devolver_comando = "so.dll -r"    '// reiniciar el equipo
  Case (3)
  devolver_comando = "so.dll -a"    '// anular el apagado del equipo
  Case (4)
  devolver_comando = "so.dll -m"    '// equipo que se / apagara / reiniciara / anulara
  Case (5)
  devolver_comando = "so.dll -t"    '// establecer el tiempo de cierre de apagado en xx segundos
  Case (6)
  devolver_comando = "so.dll -c"    '// le puedes aplicar comentarios
  Case (7)
  devolver_comando = "so.dll -f"    '// fuerza el cierre de aplicaciones sin advertir
 End Select
no_se:
End Sub

Private Sub cob1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub mc_Click(Index As Integer)
 txtd.Text = mc.Item(Index).Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmfunciones
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_controles
 On Error GoTo no_se
 If txtd.Text = "" Then
 txtd.Text = sistema.comentario
 End If
no_se:
 cargar_lenguage ' cargar lenguage
 cmdCargarComentarios
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub txtd_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(171)
 Label1.Caption = Lenguage.lenguaje_Menu(172)
 cmdcomentarios.Caption = Lenguage.lenguaje_Menu(173)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(186)
 cmdAplicar.Caption = Lenguage.lenguaje_Menu(187)
 labinfo.Caption = Lenguage.lenguaje_Menu(188)
 lbld.Caption = Lenguage.lenguaje_Menu(189)
End Sub
Private Sub cmdCargarComentarios()
On Error GoTo nose
Dim cargar As String
Dim r As Integer
r = 1
Open "comentarios.txt" For Input As 1
 Do While Not EOF(1)
       Line Input #1, cargar
       If mc(0).Caption = "" Then
          mc(0).Caption = cargar
       Else
       Load mc(r)
       mc(r).Caption = cargar
       mc(r).Visible = True
       r = r + 1
       End If
       Loop
       Close #1
       r = 0
nose:
End Sub
