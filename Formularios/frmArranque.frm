VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmArranque 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de rutas de archivos"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "frmArranque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdcargar 
         Caption         =   "&Cargar"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdborrar 
         Caption         =   "&Borrar Selección"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton cmdborrartodo 
         Caption         =   "&Borrar Todo"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdusar 
         Caption         =   "&Usar Archivo"
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   1980
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   7575
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   -120
         Picture         =   "frmArranque.frx":0CCA
         ScaleHeight     =   420
         ScaleWidth      =   8130
         TabIndex        =   1
         Top             =   0
         Width           =   8160
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "&Historial de Archivos definidos"
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
            TabIndex        =   2
            Top             =   120
            Width           =   5175
         End
      End
      Begin MSComDlg.CommonDialog cdgAbrir 
         Left            =   5040
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Historial de Archivos definidos"
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   7515
      End
   End
End
Attribute VB_Name = "frmArranque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Iniciar Archivo con el  programa Virtual Martin temporize v1.7
'* Historial de Rutas de Archivo
'*
'***************************************************************************
Private Sub cmdAceptar_Click()
 externosF.guardar_Archivo_Externo
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdborrar_Click()
If Not (List1.ListIndex = -1) Then
 Select Case MsgBox("Quieres eliminar este Archivo definido de la Lista" _
 , vbYesNo + vbInformation)
  Case (vbYes)
   List1.RemoveItem (List1.ListIndex)
 End Select
End If
End Sub

Private Sub cmdborrar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdborrartodo_Click()
 Select Case MsgBox("Quieres eliminar todos los Archivos definidos en el Historial" _
 , vbYesNo + vbInformation)
  Case (vbYes)
  List1.Clear
 End Select
End Sub

Private Sub cmdborrartodo_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdcargar_Click()
With cdgAbrir
 If .CancelError = False Then
 .DialogTitle = "Virtual Martin temporize v1.0: Cargar Archivo"
 .Filter = "Virtual Martin temporize v1.0 (*.vmt)|*.vmt|todos los Archivos (*.*)|*.*|"
 .ShowOpen
 If .FileName = "" Then
 MsgBox "Tienes que seleccionar un Archivo para poder Abrirlo", vbInformation
 End If
 If .FileName <> "" Then
 List1.AddItem .FileName
 End If
 End If
End With
End Sub

Private Sub cmdcargar_KeyPress(KeyAscii As Integer)
salir_op KeyAscii
End Sub

Private Sub cmdusar_Click()
If cmdusar.Caption = Lenguage.lenguaje_Menu(126) Then
 If Not (List1.ListIndex = -1) Then
 MsgBox Lenguage.lenguaje_Menu(129) & "" & List1.List(List1.ListIndex)
 externosF.xselecionado = List1.List(List1.ListIndex)
 externosF.guardar_selecionado
 End If
 cmdusar.Caption = Lenguage.lenguaje_Menu(127)
 ElseIf cmdusar.Caption = Lenguage.lenguaje_Menu(127) Then
 Select Case MsgBox(Lenguage.lenguaje_Menu(130), vbYesNo + vbInformation)
  Case (vbYes)
   externosF.xselecionado = ""
   externosF.guardar_selecionado
   End Select
   cmdusar.Caption = Lenguage.lenguaje_Menu(126)
 End If
End Sub

Private Sub cmdusar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()

 Me.Icon = frmprograma.Icon
 externosF.Abrir_Archivo_Externo
 cargar_lenguage ' carga el lenguage
 Label2.Caption = Label1.Caption
 
  'carga Skins con el recurso del formulario requerido
 cargar_Skins Me


End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmArranque
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(120)
 Label1.Caption = Lenguage.lenguaje_Menu(121)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(122)
 cmdcargar.Caption = Lenguage.lenguaje_Menu(123)
 cmdborrar.Caption = Lenguage.lenguaje_Menu(124)
 cmdborrartodo.Caption = Lenguage.lenguaje_Menu(125)
 cmdusar.Caption = Lenguage.lenguaje_Menu(126)
 cmdaceptar.Caption = Lenguage.lenguaje_Menu(128)
End Sub
