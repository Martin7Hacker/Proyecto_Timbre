VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmhistorial 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial"
   ClientHeight    =   6495
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "frmhistorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9675
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   255
      BackColor       =   14585656
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmhistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Historial para Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim d As Long

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_datos
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Form_Resize()
 On Error GoTo no_se
 ListView1.Width = Me.Width - 400
 ListView1.Height = Me.Height - 1400
no_se:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmprograma.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdguardarysalir_Click()
 frmprograma.guardard_Click
 End
End Sub

Private Sub cmdguardarysalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 End
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_datos()
 Const espacio As String = "                               "
 On Error GoTo no_se
  With frmprograma
  Dim ah As Integer
  Dim v As String
  Dim et As ListItem
  With ListView1.ColumnHeaders
  .Add , , "A�o"
  .Add , , "Mes"
  .Add , , "Dia"
  .Add , , "Hora"
  .Add , , "Tipo"
  .Add , , "Segundos"
  .Add , , "Comentario"
 End With
 With ListView1
 .View = lvwReport
 .LabelEdit = lvwManual
 ' Permitir m�ltiple selecci�n
 .MultiSelect = False
 ' Para que al perder el foco,
 ' se siga viendo el que est� seleccionado
 .HideSelection = False
 End With
 ListView1.View = lvwReport
 For ah = 0 To .listado(0).ListCount - 1
 If .listado(1).List(ah) = "Salida" Then
 v = "   "
 Else
 v = ""
 End If
 d = Int(ah) + 1
 With ListView1.ListItems.Add(, , "Evento_____ " & d)
 .SubItems(1) = frmprograma.listado(0).List(ah)
 .SubItems(2) = frmprograma.listado(1).List(ah)
 .SubItems(3) = "seg. " & frmprograma.listado(2).List(ah)
 .SubItems(4) = frmprograma.listado(3).List(ah)
 '  List1.AddItem  & espacio & .listado(1).List(ah) & v & espacio & .listado(2).List(ah) & espacio & .listado(3).List(ah)
 End With
 Next ah
End With
no_se:
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmVisorEventos
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub imprimir_Click()
 ModImprimir.Imprimir_ListView
End Sub

Private Sub imprimirMas_Click()
 frmimpresor.Show 1
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
 Select Case Button
  Case (2)
  PopupMenu menu
 End Select
End Sub

