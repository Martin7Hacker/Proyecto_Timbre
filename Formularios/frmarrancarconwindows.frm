VERSION 5.00
Begin VB.Form frmarrancarconwindows 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Martin temporize: Iniciar con Windows"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmarrancarconwindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   -480
         Picture         =   "frmarrancarconwindows.frx":0CCA
         ScaleHeight     =   420
         ScaleWidth      =   8130
         TabIndex        =   1
         Top             =   0
         Width           =   8160
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "&Iniciar con Windows Automaticamente"
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
            Left            =   720
            TabIndex        =   2
            Top             =   120
            Width           =   5175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EDAC85&
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   5175
         Begin VB.CommandButton cmdArranciar 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdnoarrancar 
            Caption         =   "&No"
            Height          =   495
            Left            =   4080
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdaplicar 
            Caption         =   "&Si"
            Height          =   495
            Left            =   3000
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "� Iniciar programa con el Sistema Operativo Windows ?"
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
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   4725
      End
   End
End
Attribute VB_Name = "frmarrancarconwindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Iniciar con Windows para el programa Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Option Explicit
'Constantes de la Rama del registro para los path de _
 las aplicaciones que inician con Windows
Const RAMA_RUN_WINDOWS As String = "SOFTWARE\Microsoft\" & _
                                   "Windows\CurrentVersion\Run"
Private Sub cmdAplicar_Click()
frmprograma.Enabled = True
Unload Me
End Sub

Private Sub cmdaplicar_KeyPress(KeyAscii As Integer)
salir_op KeyAscii
End Sub

Private Sub cmdArranciar_Click()
 Dim Path_Programa, _
 Titulo_Programa As String
 Dim Ret As Boolean
  On Error GoTo nose
    Path_Programa = App.Path & "\" & App.EXEName & ".exe"
    Titulo_Programa = App.Title
    Ret = EstablecerValor(HKEY_LOCAL_MACHINE1, _
                    RAMA_RUN_WINDOWS, _
                    Titulo_Programa, _
                    Path_Programa, REG_SZ1)
'si retorna True es por que cre� el dato correctamente
    If Ret Then
       MsgBox Lenguage.lenguaje_Menu(136), vbInformation
    Else
       MsgBox Lenguage.lenguaje_Menu(137), vbCritical
    End If
nose:
End Sub

Private Sub cmdArranciar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdnoarrancar_Click()
 Dim Titulo_Programa As String
 Titulo_Programa = App.Title
 Call EliminarValor(HKEY_LOCAL_MACHINE, _
                   RAMA_RUN_WINDOWS, _
                   Titulo_Programa)
                   MsgBox Lenguage.lenguaje_Menu(138), vbInformation
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmarrancarconwindows
End Sub

Private Sub cmdnoarrancar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_lenguage ' carga el lenguage del programa
 
 'carga Skins con el recurso del formulario requerido
cargar_Skins Me

End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguaje_Menu(131)
 Label2.Caption = Lenguage.lenguaje_Menu(131)
 Label1.Caption = Lenguage.lenguaje_Menu(132)
 cmdArranciar.Caption = Lenguage.lenguaje_Menu(133)
 cmdnoarrancar.Caption = Lenguage.lenguaje_Menu(134)
 cmdAplicar.Caption = Lenguage.lenguaje_Menu(135)
End Sub


