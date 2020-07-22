VERSION 5.00
Begin VB.Form frmtimbre 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evento en ejecución"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   Icon            =   "frmtimbre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EDAC85&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   8235
      TabIndex        =   15
      Top             =   120
      Width           =   8295
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   2760
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label labinfo 
         BackStyle       =   0  'Transparent
         Caption         =   "El Timbre se esta ejecutando..."
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
         TabIndex        =   16
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Timer tiempo 
      Interval        =   1000
      Left            =   600
      Top             =   5640
   End
   Begin VB.Timer timreloj 
      Interval        =   1000
      Left            =   120
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDAC85&
      Caption         =   "Comentarios"
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EDAC85&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   7995
         TabIndex        =   17
         Top             =   3720
         Width           =   8055
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo :"
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
            Index           =   4
            Left            =   0
            TabIndex        =   22
            Top             =   960
            Width           =   4485
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora que se Activo :"
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
            Index           =   3
            Left            =   0
            TabIndex        =   21
            Top             =   720
            Width           =   5370
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Timpo Restante :"
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
            Index           =   2
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   5370
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Timpo Trascurrido :"
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
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   5370
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Timpo Total :"
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
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   5370
         End
      End
      Begin VB.Frame frmsolo_hora 
         BackColor       =   &H00EDAC85&
         Height          =   735
         Left            =   6360
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label labhora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solo Hora."
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
            Left            =   150
            TabIndex        =   14
            ToolTipText     =   "Se ejecuta siempre sin importar el Día de la Semana"
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.Frame fram_dias 
         BackColor       =   &H00EDAC85&
         Height          =   2175
         Left            =   6360
         TabIndex        =   4
         ToolTipText     =   "Listado de Progrmación de los dias o el dia que queres activar el timbre."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "domingos"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Sabados"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Viernes"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Jueves"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Miercoles"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Martes"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EDAC85&
            Caption         =   "Lunes"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FF00FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DIAS"
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
            Index           =   5
            Left            =   360
            TabIndex        =   12
            Top             =   1850
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   5520
         Picture         =   "frmtimbre.frx":0CCA
         ScaleHeight     =   975
         ScaleWidth      =   2655
         TabIndex        =   2
         Top             =   240
         Width           =   2655
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cargando..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EDAC85&
            Height          =   435
            Left            =   150
            TabIndex        =   3
            Top             =   170
            Width           =   2220
         End
      End
      Begin VB.TextBox Text1 
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
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmtimbre.frx":10C80
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmtimbre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Detonador de  Martin temporize v1.7
'*
'*
'***************************************************************************
Public timpo_programado, restante, trascurrido As Integer
Public comentario_general As String

Private Sub Command1_Click()
 Finalizar
 Unload Me
End Sub

Private Sub Form_Load()
 On Error GoTo nose
 Me.Icon = frmprograma.Icon
 restante = timpo_programado
 Text1.Text = comentario_general
 puertof.disparar_bit ' Enciendo el Timbre
 Call cargarIdioma


'carga Skins con el recurso del formulario requerido
cargar_Skins Me

nose:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error GoTo no_se
 Shell frmprograma.liscomando.List(frmprograma.liscomando.ListIndex), _
 vbNormalNoFocus
no_se:
 Command1_Click
 timpo_programado = 0
 restante = 0
 trascurrido = 0
 'Finalizar no dispara al puerto
 frmprograma.guardard_Click
End Sub

Public Sub Finalizar()
puertof.apagar_puertos ' apaga todos los puertos de LTP1
End Sub


Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub tiempo_Timer()
 trascurrido = trascurrido + 1: restante = restante - 1
 Label1(1).Caption = lenguaje_Menu(342) & " " & trascurrido & " " & lenguaje_Menu(382)
 Label1(2).Caption = lenguaje_Menu(343) & " " & restante & " " & lenguaje_Menu(382)
 Command1.Caption = lenguaje_Menu(346) & " " & "(" & restante & ")"
 funcin_cerrar
End Sub

Private Sub timreloj_Timer()
 Label2.Caption = Time & " " & lenguaje_Menu(350)
End Sub

Private Sub funcin_cerrar()
 If timpo_programado = trascurrido Then
 trascurrido = 0 ' destrullo la hora
 fram_dias.Visible = False
 frmsolo_hora.Visible = False
 'apagarTodo_Puerto
 Unload Me
 End If
End Sub
Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(328)
labinfo.Caption = lenguaje_Menu(329)
Frame1.Caption = lenguaje_Menu(330)
Label2.Caption = lenguaje_Menu(331)
Check1(0).Caption = lenguaje_Menu(332)
Check1(1).Caption = lenguaje_Menu(333)
Check1(2).Caption = lenguaje_Menu(334)
Check1(3).Caption = lenguaje_Menu(335)
Check1(4).Caption = lenguaje_Menu(336)
Check1(5).Caption = lenguaje_Menu(337)
Check1(6).Caption = lenguaje_Menu(338)
Label1(5).Caption = lenguaje_Menu(339)
labhora.Caption = lenguaje_Menu(340)
Label1(0).Caption = lenguaje_Menu(341)
Label1(1).Caption = lenguaje_Menu(342)
Label1(2).Caption = lenguaje_Menu(343)
Label1(3).Caption = lenguaje_Menu(344)
Label1(4).Caption = lenguaje_Menu(345)
Command1.Caption = lenguaje_Menu(346)
End Sub
