VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPrincipioFinal 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdaplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2200
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EDAC85&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   8715
      TabIndex        =   4
      Top             =   1080
      Width           =   8775
      Begin VB.Label Label1 
         BackColor       =   &H00EDAC85&
         Caption         =   "Final:"
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
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   8715
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      Begin VB.Label Label2 
         BackColor       =   &H00EDAC85&
         BackStyle       =   0  'Transparent
         Caption         =   "Principio:"
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
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   8655
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1191
      _Version        =   327682
      BorderStyle     =   1
      SelectRange     =   -1  'True
      TickStyle       =   1
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1191
      _Version        =   327682
      BorderStyle     =   1
      SelectRange     =   -1  'True
   End
End
Attribute VB_Name = "frmPrincipioFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAplicar_Click()
  fc.principio = Me.Slider1.Value
  fc.final = Me.Slider2.Value
  frmentradasalida.cmdmod.Enabled = True
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub



Private Sub Form_Load()
Dim reg As Long
  Me.Icon = frmprograma.Icon
 On Error GoTo nose
 If Not (frmprograma.listado(0).ListCount - 1 = 0) Then
  reg = frmprograma.listado(0).ListCount - 1
 End If
 Me.Slider1.Min = 0
 Me.Slider1.Max = reg
 Me.Slider2.Max = Me.Slider1.Max
 Me.Slider2.Min = Me.Slider1.Min
 Me.Slider2.Value = reg
 Me.Slider1.Enabled = True
 Me.Slider2.Enabled = True
 Me.Slider1.Value = fc.principio
 Me.Slider2.Value = fc.final
 
 
nose:

'carga Skins con el recurso del formulario requerido
 cargar_Skins Me


End Sub

Private Sub Slider1_Change()
Slider1_Click
End Sub
Private Sub Slider1_Click()
Slider1.ToolTipText = Slider1.Value
End Sub
Private Sub Slider1_Scroll()
Slider1_Click
End Sub

Private Sub Slider2_Change()
Slider2_Click
End Sub
Private Sub Slider2_Click()
Slider2.ToolTipText = Slider2.Value
End Sub
Private Sub Slider2_Scroll()
Slider2_Click
End Sub

