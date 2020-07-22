VERSION 5.00
Begin VB.Form frmimpresor 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresor por cantidad"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   Icon            =   "frmimpresor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Mandar a Imprimir las Copias"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00EDAC85&
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   240
      Width           =   735
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmimpresor.frx":0CCA
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdmas 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdmenos 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00EDAC85&
         Height          =   255
         Left            =   960
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   6
         Top             =   240
         Width           =   855
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Copias:"
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
            Width           =   645
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00EDAC85&
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   240
         Width           =   495
         Begin VB.Image Image2 
            Height          =   480
            Left            =   0
            Picture         =   "frmimpresor.frx":1994
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.TextBox txtcop 
         BackColor       =   &H00EDAC85&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   720
         ScaleHeight     =   795
         ScaleWidth      =   0
         TabIndex        =   2
         Top             =   180
         Width           =   60
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmimpresor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Poder Imprimir para Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim copias As Long

Private Sub cmdAceptar_Click()
 Dim ip As Long
 For ip = 1 To copias
 ModImprimir.Imprimir_ListView
 Next
 Unload Me
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdmas_Click()
 copias = copias + 1
 txtcop.Text = copias
 cmdmenos.Enabled = True
End Sub

Private Sub cmdmenos_Click()
 If copias = 1 Then
 cmdmenos.Enabled = False
 Else
 copias = copias - 1
 End If
 txtcop.Text = copias
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 copias = 1
 txtcop.Text = copias
 cargar_lenguage 'cargar lenguage
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(190)
 Label1.Caption = Lenguage.lenguaje_Menu(191)
 cmdmenos.Caption = Lenguage.lenguaje_Menu(192)
 cmdmas.Caption = Lenguage.lenguaje_Menu(193)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(194)
 cmdaceptar.Caption = Lenguage.lenguaje_Menu(195)
End Sub
