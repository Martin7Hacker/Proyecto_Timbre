VERSION 5.00
Begin VB.Form frmpuerto 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puerto de Salida"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   ClipControls    =   0   'False
   Icon            =   "frmpuerto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdnormal 
      Caption         =   "&normal"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5475
      TabIndex        =   9
      Top             =   0
      Width           =   5535
      Begin VB.Label Labelbuerto 
         BackStyle       =   0  'Transparent
         Caption         =   "Usted Tiene que tener conocimiento antes de realizar algun cambio Aqu�."
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
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5205
      End
   End
   Begin VB.CheckBox pin8 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 8"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox pin7 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin6 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin5 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin4 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox pin3 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox pin2 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox pin1 
      BackColor       =   &H00EDAC85&
      Caption         =   "Pin 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Caption         =   "&Salida 5v"
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
      Height          =   855
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   $"frmpuerto.frx":0CCA
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "frmpuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Conexi�n por Puerto Paralelo de Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdCancelar_Click()
 cerrar
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdnormal_Click()
 pin1.Value = 1
 pin2.Value = 0
 pin3.Value = 1
 pin4.Value = 0
 pin5.Value = 0
 pin6.Value = 0
 pin7.Value = 0
 pin8.Value = 0
 almacenar_datos 'llamada al procedimiento
End Sub

Private Sub cmdnormal_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 cerrar
End Sub

Private Sub cerrar()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub almacenar_datos()
 puertof.pu1 = pin1.Value
 puertof.pu2 = pin2.Value
 puertof.pu3 = pin3.Value
 puertof.pu4 = pin4.Value
 puertof.pu5 = pin5.Value
 puertof.pu6 = pin6.Value
 puertof.pu7 = pin7.Value
 puertof.pu8 = pin8.Value
End Sub

Private Sub cargar_datos()
 pin1.Value = puertof.pu1
 pin2.Value = puertof.pu2
 pin3.Value = puertof.pu3
 pin4.Value = puertof.pu4
 pin5.Value = puertof.pu5
 pin6.Value = puertof.pu6
 pin7.Value = puertof.pu7
 pin8.Value = puertof.pu8
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 cargar_datos
 Me.Icon = frmprograma.Icon
 Call cargarlenguaje

'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
 almacenar_datos
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmpuerto
End Sub

Private Sub cargarlenguaje()
 Me.Caption = lenguaje_Menu(352)
 Labelbuerto.Caption = lenguaje_Menu(353)
 Frame1.Caption = lenguaje_Menu(354)
  pin1.Caption = lenguaje_Menu(355)
  pin2.Caption = lenguaje_Menu(356)
  pin3.Caption = lenguaje_Menu(357)
  pin4.Caption = lenguaje_Menu(358)
  pin5.Caption = lenguaje_Menu(359)
  pin6.Caption = lenguaje_Menu(360)
  pin7.Caption = lenguaje_Menu(361)
  pin8.Caption = lenguaje_Menu(362)
  cmdCancelar.Caption = lenguaje_Menu(363)
  cmdnormal.Caption = lenguaje_Menu(364)
  cmdSalir.Caption = lenguaje_Menu(365)
End Sub
