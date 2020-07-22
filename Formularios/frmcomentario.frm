VERSION 5.00
Begin VB.Form frmcomentario 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir comentarios"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCargarComentarios 
      Caption         =   "&Cargar "
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdAniadir 
      Caption         =   "&Añadir"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdEliminarTodo 
      Caption         =   "&Eliminar Todo"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdeliminarselecionado 
      Caption         =   "&Eliminar Seleciónado"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ListBox lstComentario 
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
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8655
   End
   Begin VB.TextBox txtComentario 
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
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      Begin VB.Label lblComentario 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   50
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmcomentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* comentarios en  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdAniadir_Click()
If Not (txtComentario.Text = "") Then
lstComentario.AddItem txtComentario.Text
End If
txtComentario.Text = ""
End Sub

Private Sub cmdAplicar_Click()
Dim r As Integer
Open "comentarios.txt" For Output As 1
 For r = 0 To lstComentario.ListCount - 1
 Print #1, lstComentario.List(r)
 Next r
Close #1
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdCargarComentarios_Click()
lstComentario.Clear
Dim cargar As String

Open "comentarios.txt" For Input As 1
 Do While Not EOF(1)
       Line Input #1, cargar
       lstComentario.AddItem cargar
       Loop
       Close #1
End Sub

Private Sub cmdeliminarselecionado_Click()
If Not (lstComentario.ListIndex = -1) Then
 Select Case MsgBox(lenguaje_Menu(273) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.RemoveItem (lstComentario.ListIndex)
 End Select
End If
End Sub

Private Sub cmdEliminarTodo_Click()
If Not (lstComentario.ListIndex <= -1) Then
 Select Case MsgBox(lenguaje_Menu(274) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.Clear
 End Select
End If
End Sub

Private Sub Form_Load()
Me.Icon = frmprograma.Icon
cmdCargarComentarios_Click
cargarIdioma

'carga Skins con el recurso del formulario requerido
cargar_Skins Me
Picture1.BackColor = Me.BackColor

End Sub
Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(265)
lblComentario.Caption = lenguaje_Menu(266)
cmdAniadir.Caption = lenguaje_Menu(267)
cmdCargarComentarios.Caption = lenguaje_Menu(268)
cmdCancelar.Caption = lenguaje_Menu(269)
cmdeliminarselecionado.Caption = lenguaje_Menu(270)
cmdEliminarTodo.Caption = lenguaje_Menu(271)
cmdAplicar.Caption = lenguaje_Menu(272)
End Sub

Private Sub lstComentario_Click()
txtComentario.Text = lstComentario.List(lstComentario.ListIndex)
End Sub

Private Sub lstComentario_Scroll()
lstComentario_Click
End Sub
