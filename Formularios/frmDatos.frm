VERSION 5.00
Begin VB.Form frmDatos 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personalizar"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   Icon            =   "frmDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   43
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   2760
      TabIndex        =   42
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   6720
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   12
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   5955
      TabIndex        =   39
      Top             =   5040
      Width           =   6015
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   12
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   11
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   37
      Top             =   4680
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   10
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   35
      Top             =   4320
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   9
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   33
      Top             =   3960
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   8
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   31
      Top             =   3600
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   7
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   29
      Top             =   3240
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   6
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   27
      Top             =   2880
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   5
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   25
      Top             =   2520
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   4
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   3
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   1800
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   2
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   1440
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   1
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   15
      Top             =   720
      Width           =   1815
      Begin VB.Label labdatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
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
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "frmDatos.frx":0CCA
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   13
      Top             =   0
      Width           =   8160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Perzonalizar Datos"
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
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   12
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5400
      Width           =   6015
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   11
      Left            =   2040
      TabIndex        =   11
      Top             =   4680
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   2040
      TabIndex        =   10
      Top             =   4320
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   2040
      TabIndex        =   9
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDAC85&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Datos de los creadores de Timbres en Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdAceptar_Click()
 guardar
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub guardar()
 abrirF.xnombre = txtdato(0).Text
 abrirF.xnombre2 = txtdato(1).Text
 abrirF.xapellido = txtdato(2).Text
 abrirF.xapellido2 = txtdato(3).Text
 abrirF.xdireccion = txtdato(4).Text
 abrirF.xdireccion2 = txtdato(5).Text
 abrirF.xlocalidad = txtdato(6).Text
 abrirF.xPais = txtdato(7).Text
 abrirF.xtelefono = txtdato(8).Text
 abrirF.xcel = txtdato(9).Text
 abrirF.xemail = txtdato(10).Text
 abrirF.xfacebook = txtdato(11).Text
 abrirF.xcomentario_general = txtdato(12).Text
 MsgBox Lenguage.lenguaje_Menu(170), vbInformation
End Sub

Private Sub mostrar()
 txtdato(0).Text = abrirF.xnombre
 txtdato(1).Text = abrirF.xnombre2
 txtdato(2).Text = abrirF.xapellido
 txtdato(3).Text = abrirF.xapellido2
 txtdato(4).Text = abrirF.xdireccion
 txtdato(5).Text = abrirF.xdireccion2
 txtdato(6).Text = abrirF.xlocalidad
 txtdato(7).Text = abrirF.xPais
 txtdato(8).Text = abrirF.xtelefono
 txtdato(9).Text = abrirF.xcel
 txtdato(10).Text = abrirF.xemail
 txtdato(11).Text = abrirF.xfacebook
 txtdato(12).Text = abrirF.xcomentario_general
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdlimpiar_Click()
 Select Case MsgBox(Lenguage.lenguaje_Menu(169) _
 , vbYesNo + vbInformation)
 Case (vbYes)
  Dim l As Byte
  For l = 0 To 12
   txtdato(l).Text = ""
  Next
 End Select
End Sub

Private Sub cmdlimpiar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 mostrar
 Me.Icon = frmprograma.Icon
 cargar_lenguage ' cargar lenguage
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me
 

End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmDatos
End Sub




Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(150)
 Label1.Caption = Lenguage.lenguaje_Menu(151)
 cmdaceptar.Caption = Lenguage.lenguaje_Menu(152)
 labdatos(0).Caption = Lenguage.lenguaje_Menu(153)
 labdatos(1).Caption = Lenguage.lenguaje_Menu(154)
 labdatos(2).Caption = Lenguage.lenguaje_Menu(155)
 labdatos(3).Caption = Lenguage.lenguaje_Menu(156)
 labdatos(4).Caption = Lenguage.lenguaje_Menu(157)
 labdatos(5).Caption = Lenguage.lenguaje_Menu(158)
 labdatos(6).Caption = Lenguage.lenguaje_Menu(159)
 labdatos(7).Caption = Lenguage.lenguaje_Menu(160)
 labdatos(8).Caption = Lenguage.lenguaje_Menu(161)
 labdatos(9).Caption = Lenguage.lenguaje_Menu(162)
 labdatos(10).Caption = Lenguage.lenguaje_Menu(163)
 labdatos(11).Caption = Lenguage.lenguaje_Menu(164)
 labdatos(12).Caption = Lenguage.lenguaje_Menu(165)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(166)
 cmdlimpiar.Caption = Lenguage.lenguaje_Menu(167)
 cmdaceptar.Caption = Lenguage.lenguaje_Menu(168)
End Sub

Private Sub Picture1_KeyPress(Index As Integer, KeyAscii As Integer)
salir_op KeyAscii
End Sub
