VERSION 5.00
Begin VB.Form frmutilizarManual 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilizar Timbre Manualmente"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "frmutilizarManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdapagado 
      Caption         =   "&Apagado"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdencendido 
      Caption         =   "&Encendido"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "frmutilizarManual.frx":0CCA
      ScaleHeight     =   735
      ScaleWidth      =   3615
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000C0&
         BorderWidth     =   10
         Height          =   375
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   165
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "Led Que Muestra el Estado del Timbre ."
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   -600
      ScaleHeight     =   0
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmutilizarManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Detonador Manual de Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdAceptar_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdapagado_Click()
 On Error GoTo no_se
 Shape1.BackColor = &H80FF80
 Shape1.BorderColor = &HC000&
 Label1.Caption = lenguaje_Menu(324)
 frmtimbre.Finalizar ' Apaga todos los Puertos
no_se:
End Sub

Private Sub cmdapagado_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdencendido_Click()
 On Error GoTo no_se
 Shape1.BackColor = &HFF&
 Shape1.BorderColor = &HC0&
 Label1.Caption = lenguaje_Menu(325)
 puertof.disparar_bit ' Enciendo el Timbre
no_se:
End Sub

Private Sub cmdencendido_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cmdapagado_Click
 Call cargarIdioma
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
 cmdapagado_Click
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmutilizarManual
End Sub

Private Sub cargarIdioma()
 Me.Caption = lenguaje_Menu(321)
 cmdencendido.Caption = lenguaje_Menu(322)
 cmdapagado.Caption = lenguaje_Menu(323)
 Label1.Caption = lenguaje_Menu(324)
 Label2.ToolTipText = lenguaje_Menu(327)
End Sub
