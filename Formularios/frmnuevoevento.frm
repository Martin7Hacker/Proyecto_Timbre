VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmnuevoevento 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   Icon            =   "frmnuevoevento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EDAC85&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   3915
      TabIndex        =   19
      Top             =   0
      Width           =   3975
      Begin VB.Label labinfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Nuevo Evento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FBF3E8&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   1950
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FBF3E8&
         BorderWidth     =   2
         X1              =   120
         X2              =   2520
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.CommandButton cmdfunct 
      Caption         =   "&Funciones al sistema"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdcomentarios 
      Caption         =   "&comentarios:"
      Height          =   315
      Left            =   5640
      TabIndex        =   17
      Top             =   650
      Width           =   1335
   End
   Begin VB.CommandButton boton 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton boton 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FBF3E8&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         ItemData        =   "frmnuevoevento.frx":0CCA
         Left            =   840
         List            =   "frmnuevoevento.frx":0CCC
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FBF3E8&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmnuevoevento.frx":0CCE
         Left            =   840
         List            =   "frmnuevoevento.frx":0CD0
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1050
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FBF3E8&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmnuevoevento.frx":0CD2
         Left            =   600
         List            =   "frmnuevoevento.frx":0CD4
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   295
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   -2147483647
         Format          =   56885250
         UpDown          =   -1  'True
         CurrentDate     =   0.805555555555556
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EDAC85&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   2355
         TabIndex        =   21
         Top             =   240
         Width           =   2415
         Begin VB.Label etiqueta 
            BackStyle       =   0  'Transparent
            Caption         =   "Filtro :"
            ForeColor       =   &H00FBF3E8&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   420
         End
         Begin VB.Label etiqueta 
            BackStyle       =   0  'Transparent
            Caption         =   "Intervalo :"
            ForeColor       =   &H00FBF3E8&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   840
            Width           =   705
         End
         Begin VB.Label etiqueta 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo :"
            ForeColor       =   &H00FBF3E8&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   360
            Width           =   600
         End
         Begin VB.Label etiqueta 
            BackStyle       =   0  'Transparent
            Caption         =   "Hora :"
            ForeColor       =   &H00FBF3E8&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FBF3E8&
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Domingos."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Sabado ."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Viernes ."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Jueves ."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Miercoles."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Martes ."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Lunes ."
         ForeColor       =   &H00FBF3E8&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario :"
         ForeColor       =   &H00FBF3E8&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   14
      Top             =   0
      Width           =   0
   End
   Begin HookMenu.XpMenu XpMenu12 
      Left            =   4200
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
      Bmp:1           =   "frmnuevoevento.frx":0CD6
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
   Begin VB.Menu comentar 
      Caption         =   "menú"
      Visible         =   0   'False
      Begin VB.Menu mc 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmnuevoevento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Nuevo Evento y Modificación para Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim nuevoEvento As evento

Private Sub boton_Click(Index As Integer)
 frmprograma.Enabled = True
 With frmprograma
 Select Case Index
  Case (0)
  Unload Me
  Case (1)
 If boton(1).Caption = Lenguage.lenguaje_Menu(224) Then
 nuevo_evento_de_dias
 Crear ' crea un nunevo evento de timbre
 sistema.ingresarDatos
 ElseIf boton(1).Caption = Lenguage.lenguaje_Menu(225) Then
 'selección
 Select Case MsgBox(Lenguage.lenguaje_Menu(231) _
 , vbYesNo + vbInformation, lenguaje_Menu(8))
  Case (vbYes)

  labinfo.Caption = Lenguage.lenguaje_Menu(226)
  
  .listado(0).List(.listado(0).ListIndex) = DTPicker1.Value
  .listado(1).List(.listado(1).ListIndex) = Combo1(0).Text
  .listado(2).List(.listado(2).ListIndex) = Combo1(1).Text
  .listado(3).List(.listado(3).ListIndex) = Text1.Text
  .liscomando.List(.liscomando.ListIndex) = frmfunciones.devolver_comando
  'dias Set'
 set_dias ' cambia los dias de la semana
 .Filtro.List(.Filtro.ListIndex) = Combo1(2).ListIndex
 Unload Me
 End Select
 End If
 End Select
 End With
End Sub

Private Sub Crear()
 Set nuevoEvento = New evento
 With nuevoEvento
 .vHora.Add DTPicker1.Value
 .vTipo.Add Combo1(0).Text
 .vIntervalo.Add Combo1(1).Text
 .vtipod.Add Combo1(2).Text
 .vDescripcion.Add Text1.Text
 End With
 With frmprograma
 Dim recor As Integer
 For recor = 1 To nuevoEvento.vHora.Count
 .listado(0).AddItem nuevoEvento.vHora(recor)
 .listado(1).AddItem nuevoEvento.vTipo(recor)
 .listado(2).AddItem nuevoEvento.vIntervalo(recor)
 .listado(3).AddItem nuevoEvento.vDescripcion(recor)
 Next
 End With
 Unload Me
End Sub

Private Sub boton_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdobsiones_Click()
 PopupMenu obsiones
End Sub

Private Sub cmdobsiones_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdcomentarios_Click()
 PopupMenu comentar, , cmdcomentarios.Left, cmdcomentarios.Top
End Sub

Private Sub cmdfunct_Click()
 If boton(1).Caption = lenguaje_Menu(92) Then
 frmfunciones.cmdAplicar.Caption = lenguaje_Menu(92)
 End If
 frmfunciones.Show 1
End Sub

Private Sub cmdfunct_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Combo1_Click(Index As Integer)
 Select Case Index
  Case (2)
 Select Case Combo1(2).ListIndex
  Case (0)
  visiblex False
  activado 0
  Dim td As Byte
  For td = 0 To 6
  Check1(CInt(td)).Value = 1
  Next
 Case (1)
 visiblex True
 activado 0
 End Select
 End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
Call cmdCargarComentarios: Call cargarIdioma
 Me.Icon = frmprograma.Icon
 Combo1(2).ListIndex = CInt(MemoriaF.numero)
 visiblex CInt(MemoriaF.numero)
 DTPicker1.Value = Time
 agregar_elementos
 If MemoriaF.dias = True Then
 devolver_dias
 End If
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub agregar_elementos()
 Dim numero As Integer
 Combo1(0).ListIndex = 0
 For numero = 1 To 77
 Combo1(1).AddItem (numero)
 Next
 Combo1(1).ListIndex = 4
End Sub

Private Sub visiblex(ByVal visilblex As Boolean)
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Enabled = visilblex
 Next
End Sub

Private Sub activado(ByVal activado As Byte)
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Value = activado
 Next
End Sub

Private Sub almanaque_Click()
 frmalmanaque.Show 1
End Sub

Private Sub nuevo_evento_de_dias()
 Const nulo As String = "0"      'nulo
 Const lunes As String = "2"     'lunes
 Const martes As String = "3"    'martes
 Const miercoles As String = "4" 'miercoles
 Const jueves As String = "5"    'jueves
 Const viernes As String = "6"   'viernes
 Const sabado As String = "7"    'sabado
 Const domingo As String = "1"   'domingo
 With frmprograma
 Select Case Check1(0).Value     ' Lunes
  Case (1)
  .lunes(0).AddItem lunes
  Case (0)
  .lunes(0).AddItem nulo
 End Select
Select Case Check1(1).Value      ' Martes
 Case (1)
 .martes.AddItem martes          ' Martes
 Case (0)
 .martes.AddItem nulo
End Select
Select Case Check1(2).Value ' Miercoles
 Case (1)
 .miercoles.AddItem miercoles
 Case (0)
 .miercoles.AddItem nulo
End Select
Select Case Check1(3).Value ' Jueves
 Case (1)
 .jueves.AddItem jueves
 Case (0)
 .jueves.AddItem nulo
End Select
Select Case Check1(4).Value ' Viernes
 Case (1)
 .viernes.AddItem viernes
 Case (0)
 .viernes.AddItem nulo
End Select
Select Case Check1(5).Value ' Sabado
 Case (1)
 .sabado.AddItem sabado
 Case (0)
 .sabado.AddItem nulo
End Select
Select Case Check1(6).Value ' Domingo
 Case (1)
 .domingo.AddItem domingo
 Case (0)
 .domingo.AddItem nulo
End Select
'***************'> Asignacion de Filtro <******************'
.Filtro.AddItem Combo1(2).ListIndex
 End With
End Sub

Public Sub devolver_dias()
 Dim dev As Integer
 For dev = 0 To frmprograma.listado(0).ListCount
 With frmprograma
 'lunes
 Select Case .lunes(0).List(.lunes(0).ListIndex)
  Case (2)
  Check1(0).Value = 1
  Case (0)
  Check1(0).Value = 0
 End Select
'martes
Select Case .martes.List(.martes.ListIndex)
 Case (3)
 Check1(1).Value = 1
 Case (0)
 Check1(1).Value = 0
End Select
'miercoles
Select Case .miercoles.List(.miercoles.ListIndex)
 Case (4)
 Check1(2).Value = 1
 Case (0)
 Check1(2).Value = 0
End Select
'jueves
Select Case .jueves.List(.jueves.ListIndex)
 Case (5)
 Check1(3).Value = 1
 Case (0)
 Check1(3).Value = 0
End Select
'viernes
Select Case .viernes.List(.viernes.ListIndex)
 Case (6)
 Check1(4).Value = 1
 Case (0)
 Check1(4).Value = 0
End Select
'sabado
Select Case .sabado.List(.sabado.ListIndex)
 Case (7)
 Check1(5).Value = 1
 Case (0)
 Check1(5).Value = 0
End Select
'domingo
Select Case .domingo.List(.domingo.ListIndex)
 Case (1)
 Check1(6).Value = 1
 Case (0)
 Check1(6).Value = 0
End Select
End With
Next dev
End Sub

Private Sub set_dias()
 With frmprograma
 'lunes
 Select Case Check1(0).Value
  Case (1)
  .lunes(0).List(.lunes(0).ListIndex) = 2
  Case (0)
  .lunes(0).List(.lunes(0).ListIndex) = 0
 End Select
 'martes
 Select Case Check1(1).Value
  Case (1)
  .martes.List(.martes.ListIndex) = 3
  Case (0)
  .martes.List(.martes.ListIndex) = 0
  End Select
 'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.List(.miercoles.ListIndex) = 4
 Case (0)
 .miercoles.List(.miercoles.ListIndex) = 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.List(.jueves.ListIndex) = 5
 Case (0)
 .jueves.List(.jueves.ListIndex) = 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.List(.viernes.ListIndex) = 6
 Case (0)
 .viernes.List(.viernes.ListIndex) = 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.List(.sabado.ListIndex) = 7
 Case (0)
 .sabado.List(.sabado.ListIndex) = 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.List(.domingo.ListIndex) = 1
 Case (0)
 .domingo.List(.domingo.ListIndex) = 0
End Select
End With
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmnuevoevento
End Sub

Private Sub mc_Click(Index As Integer)
 Text1.Text = mc.Item(Index).Caption
End Sub

Private Sub cargarIdioma()
 labinfo.Caption = lenguaje_Menu(207)
 boton(1).Caption = Lenguage.lenguaje_Menu(224)
 cmdfunct.Caption = lenguaje_Menu(208)
 etiqueta(0).Caption = Lenguage.lenguaje_Menu(209)
 etiqueta(1).Caption = Lenguage.lenguaje_Menu(210)
 etiqueta(2).Caption = Lenguage.lenguaje_Menu(211)
 etiqueta(3).Caption = Lenguage.lenguaje_Menu(212)
 Check1(0).Caption = Lenguage.lenguaje_Menu(213)
 Check1(1).Caption = Lenguage.lenguaje_Menu(214)
 Check1(2).Caption = Lenguage.lenguaje_Menu(215)
 Check1(3).Caption = Lenguage.lenguaje_Menu(216)
 Check1(4).Caption = Lenguage.lenguaje_Menu(217)
 Check1(5).Caption = Lenguage.lenguaje_Menu(218)
 Check1(6).Caption = Lenguage.lenguaje_Menu(219)
 etiqueta(4).Caption = Lenguage.lenguaje_Menu(220)
 cmdcomentarios.Caption = Lenguage.lenguaje_Menu(221)
 boton(0).Caption = Lenguage.lenguaje_Menu(222)
 Combo1(0).AddItem Lenguage.lenguaje_Menu(227)
 Combo1(0).AddItem Lenguage.lenguaje_Menu(228)
 Combo1(2).AddItem Lenguage.lenguaje_Menu(229)
 Combo1(2).AddItem Lenguage.lenguaje_Menu(230)
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
       Load frmnuevoevento.mc(r)
       mc(r).Caption = cargar
       
       mc(r).Visible = True
       r = r + 1
       End If
       Loop
       Close #1
       r = 0
nose:
End Sub
