VERSION 5.00
Begin VB.Form frmcomo 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de la Ventana Principal"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   Icon            =   "frmcomo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -120
      Picture         =   "frmcomo.frx":0CCA
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   8
      Top             =   0
      Width           =   8160
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado Grafico de la Ventana"
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
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   0
      Left            =   120
      Picture         =   "frmcomo.frx":FD8C
      ScaleHeight     =   2265
      ScaleWidth      =   3030
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   1
      Left            =   3240
      Picture         =   "frmcomo.frx":2646E
      ScaleHeight     =   2265
      ScaleWidth      =   3045
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   2
      Left            =   120
      Picture         =   "frmcomo.frx":3CDAC
      ScaleHeight     =   2250
      ScaleWidth      =   3045
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   3
         Left            =   120
         Picture         =   "frmcomo.frx":53486
         ScaleHeight     =   2265
         ScaleWidth      =   3030
         TabIndex        =   4
         Top             =   240
         Width           =   3060
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EDAC85&
         Height          =   2535
         Left            =   3240
         TabIndex        =   1
         Top             =   120
         Width           =   3375
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   3075
            TabIndex        =   10
            Top             =   600
            Width           =   3135
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado:"
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
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   3060
            End
         End
         Begin VB.ComboBox cob1 
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
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C000C0&
            Height          =   195
            Left            =   720
            TabIndex        =   3
            Top             =   600
            Width           =   2460
         End
      End
   End
End
Attribute VB_Name = "frmcomo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Cófiguracion Grafica de la Ventana de  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdAplicar_Click()
 sistema.ven = cob1.ListIndex
     externosF.GuardarVentana
            cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cob1_Change()
        mostrarVentana cob1.ListIndex
        estado_ventana
End Sub

Private Sub cob1_Click()
 cob1_Change
 estado_ventana
End Sub

Private Sub cob1_Scroll()
 cob1_Change
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_controles
 cob1.ListIndex = sistema.ven
 cargar_lenguage ' cargar lenguage
 estado_ventana
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub cargar_controles()
 With cob1
  .AddItem Lenguage.lenguaje_Menu(147)
  .AddItem Lenguage.lenguaje_Menu(148)
  .AddItem Lenguage.lenguaje_Menu(149)
 End With
End Sub

Private Sub mostrarVentana(ven As Byte)
 Select Case ven
          Case 0
           pic(3).Picture = pic(1).Picture
          Case 1
           pic(3).Picture = pic(0).Picture
          Case 2
           pic(3).Picture = pic(2).Picture
 End Select
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(143)
 Label3.Caption = Lenguage.lenguaje_Menu(143)
 Label1.Caption = Lenguage.lenguaje_Menu(144)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(145)
 cmdAplicar.Caption = Lenguage.lenguaje_Menu(146)
End Sub

Private Sub estado_ventana()
 Label2.Caption = cob1.List(cob1.ListIndex)
End Sub
