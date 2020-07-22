VERSION 5.00
Begin VB.Form frmidioma 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instancia de Idioma"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   480
      Width           =   1095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   120
      Width           =   1095
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdAplicar 
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdguardar 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdcargar 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdrenombrar 
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCargarIdioma 
      Caption         =   "..."
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2790
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdidioma 
      BackColor       =   &H00FBF3E8&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7815
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   6255
   End
   Begin VB.TextBox txtvalor 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   450
      Width           =   5175
   End
   Begin VB.ListBox lstidioma 
      Appearance      =   0  'Flat
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
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
   End
End
Attribute VB_Name = "frmidioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Configurar idioma en  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim r As Integer
Private Sub cmdAplicar_Click()
Dim int_c As Integer
Dim cargar As String
lstidioma.Clear
Open txtnombre.Text For Input As 1
 Do While Not EOF(1)
  
       Line Input #1, cargar
       Lenguage.lenguaje_Menu(int_c) = cargar
       lstidioma.AddItem cargar
       int_c = int_c + 1
       Loop
       Close #1
       int_c = 0
       Lenguage.definir_lenguage_opciones
       frmprograma.cargarIdioma
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdcargar_Click()
Dim cargar As String
lstidioma.Clear
Open txtnombre.Text For Input As 1
 Do While Not EOF(1)
  
       Line Input #1, cargar
       lstidioma.AddItem cargar
       Loop
       Close #1
End Sub

Private Sub cmdCargarIdioma_Click()
frmCargarIdioma.Show 1
End Sub

Private Sub cmdguardar_Click()
Open txtnombre.Text For Output As 1
 For r = 0 To 385
 Lenguage.lenguaje_Menu(r) = lstidioma.List(r)
 Print #1, Lenguage.lenguaje_Menu(r)
 Next r
Close #1
End Sub

Private Sub cmdrenombrar_Click()
lstidioma.List(lstidioma.ListIndex) = txtvalor.Text
End Sub

Private Sub Form_Load()
Me.Icon = frmprograma.Icon
For r = 0 To 385
lstidioma.AddItem Lenguage.lenguaje_Menu(r)
 Next r
 Call cargarIdioma
 VScroll1_Scroll
   txtnombre.Text = Lenguage.sel
   
   'carga Skins con el recurso del formulario requerido
cargar_Skins Me
   
End Sub

Private Sub cargarPrograma()
Me.Icon = frmprograma.Icon
End Sub

Private Sub lstidioma_Click()
lstidioma_Scroll
VScroll1.Value = lstidioma.ListIndex
End Sub

Private Sub lstidioma_Scroll()
txtvalor.Text = lstidioma.List(lstidioma.ListIndex)
VScroll1.Value = lstidioma.ListIndex

End Sub

Private Sub VScroll1_Change()
 On Error GoTo nose
 With VScroll1
 .Max = lstidioma.ListCount - 1
 .Min = 0
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 End With
nose:
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(232)
Label2.Caption = lenguaje_Menu(233)
Label1.Caption = lenguaje_Menu(234)
cmdrenombrar.Caption = lenguaje_Menu(235)
cmdidioma.Caption = lenguaje_Menu(236)
cmdCancelar.Caption = lenguaje_Menu(237)
cmdcargar.Caption = lenguaje_Menu(238)
cmdguardar.Caption = lenguaje_Menu(239)
cmdAplicar.Caption = lenguaje_Menu(240)
End Sub
