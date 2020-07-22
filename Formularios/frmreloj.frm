VERSION 5.00
Begin VB.Form frmreloj 
   BackColor       =   &H00DE8F38&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Virtual Martin Temporize:  Reloj Digital"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   Icon            =   "frmreloj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdaceptar 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DE8F38&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00DE8F38&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   3795
         TabIndex        =   6
         Top             =   240
         Width           =   3855
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reloj del Sistema."
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
            TabIndex        =   7
            Top             =   0
            Width           =   1545
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Picture         =   "frmreloj.frx":57E2
         ScaleHeight     =   975
         ScaleWidth      =   3855
         TabIndex        =   2
         Top             =   480
         Width           =   3855
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   75
         End
         Begin VB.Label lab_reloj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   3855
         End
      End
      Begin VB.Label labdata 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1245
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmreloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Reloj de Virtual Martin temporize v1.7
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

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 
 Call cargarIdioma
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me
 
 
End Sub

Private Sub Timer1_Timer()
 lab_reloj.Caption = Time & " " & lenguaje_Menu(350)
 Label2.Caption = lenguaje_Menu(349) & " " & Date
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmreloj
End Sub

Private Sub cargarIdioma()
  Me.Caption = lenguaje_Menu(347)
  Label1.Caption = lenguaje_Menu(348)
  cmdaceptar.Caption = lenguaje_Menu(351)
End Sub
