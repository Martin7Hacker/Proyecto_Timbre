VERSION 5.00
Begin VB.Form frmDonativos 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donativos"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "frmDonativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdcolaborar 
      Caption         =   "&Colaborar"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.Label lblcard 
         BackStyle       =   0  'Transparent
         Caption         =   "Con tarjetas de créditos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2985
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "con cuenta propia..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "para cumplir mi sueño de ir a EE:UU"
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
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7125
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Amo mucho a EE:UU"
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
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   7125
      End
   End
   Begin VB.PictureBox pdonar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00EDAC85&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1920
      MouseIcon       =   "frmDonativos.frx":0CCA
      Picture         =   "frmDonativos.frx":0FD4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox ptargeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      Picture         =   "frmDonativos.frx":155E
      ScaleHeight     =   225
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmDonativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Para realizar donacíones para el proyecto Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAceptar_Click()
 Unload Me
End Sub

Private Sub cmdcolaborar_Click()
 ptargeta_Click
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 Call cargarIdioma

'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub Label1_Click()
 ptargeta_Click
End Sub

Private Sub lblcard_Click()
 ptargeta_Click
End Sub

Private Sub pdonar_Click()
 ptargeta_Click
End Sub

Private Sub ptargeta_Click()
 Dim X As String
 X = ShellExecute(Me.hwnd, "Open" _
 , "http://martinsoft0.blogspot.com/p/donar.html", _
 &O0, &O0, 0)
 Unload Me
End Sub
Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(310)
Label2.Caption = lenguaje_Menu(311)
Label3.Caption = lenguaje_Menu(312)
Label1.Caption = lenguaje_Menu(313)
lblcard.Caption = lenguaje_Menu(314)
cmdcolaborar.Caption = lenguaje_Menu(315)
cmdaceptar.Caption = lenguaje_Menu(316)
End Sub

