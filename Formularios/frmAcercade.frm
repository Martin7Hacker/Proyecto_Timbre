VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAcercade 
   BackColor       =   &H00EDAC85&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   6045
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5850
   ClipControls    =   0   'False
   Icon            =   "frmAcercade.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4172.366
   ScaleMode       =   0  'User
   ScaleWidth      =   5493.453
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDAC85&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&Info. del sistema..."
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton cmdFacebook 
         Caption         =   "Facebook"
         Height          =   855
         Left            =   2760
         TabIndex        =   12
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton cmdTwitter 
         Caption         =   "Instagram"
         Height          =   855
         Left            =   1800
         TabIndex        =   11
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton cmdyoutube 
         Caption         =   "YouTube"
         Height          =   855
         Left            =   840
         TabIndex        =   10
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   750
         Index           =   0
         Left            =   120
         Picture         =   "frmAcercade.frx":0CCA
         ScaleHeight     =   526.75
         ScaleMode       =   0  'User
         ScaleWidth      =   526.75
         TabIndex        =   4
         ToolTipText     =   "Autor del Programa  Martin Grasso Castrillo ."
         Top             =   120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picIcon 
         BackColor       =   &H00EDAC85&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   750
         Index           =   1
         Left            =   120
         Picture         =   "frmAcercade.frx":2ABC
         ScaleHeight     =   526.75
         ScaleMode       =   0  'User
         ScaleWidth      =   526.75
         TabIndex        =   3
         ToolTipText     =   "Autor del Programa  Martin Grasso Castrillo ."
         Top             =   120
         Width           =   750
      End
      Begin VB.PictureBox picsoft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1680
         Picture         =   "frmAcercade.frx":3786
         ScaleHeight     =   315
         ScaleWidth      =   2235
         TabIndex        =   2
         Top             =   4200
         Width           =   2265
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   0
         BackColor       =   16511976
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAcercade.frx":5C88
         ForeColor       =   &H0080C0FF&
         Height          =   1305
         Left            =   840
         TabIndex        =   8
         Top             =   2040
         Width           =   3870
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Versi�n"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   7
         Top             =   660
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "T�tulo de la aplicaci�n"
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   1080
         TabIndex        =   6
         Top             =   120
         Width           =   3885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FBF3E8&
         X1              =   5640
         X2              =   120
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Caption         =   "Compilado: Canelones Tala Software."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   375
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Acerca de  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Option Explicit
Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim d As Integer
' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' N�mero de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdCambiar_Click()
 If lblDisclaimer.Visible = True Then
    picIcon.Item(0).Visible = True
    picIcon.Item(1).Visible = False
    lblDisclaimer.Visible = False
    picsoft.Visible = False
    ListView1.Visible = True
    cmdCambiar.ToolTipText = lenguaje_Menu(277)
    cmdCambiar.Caption = "3"
    ElseIf lblDisclaimer.Visible = False Then
    lblDisclaimer.Visible = True
     picsoft.Visible = True
    picIcon.Item(0).Visible = False
    picIcon.Item(1).Visible = True
    ListView1.Visible = False
    cmdCambiar.ToolTipText = lenguaje_Menu(278)
    cmdCambiar.Caption = "4"
 End If
End Sub

Private Sub cmdFacebook_Click()
 Dim X As String
 X = ShellExecute(Me.hwnd, "Open" _
 , "https://www.facebook.com/hacker.martin0", _
 &O0, &O0, 0)
 Unload Me
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub



Private Sub cmdTwitter_Click()
Dim X As String
 X = ShellExecute(Me.hwnd, "Open" _
 , "https://www.instagram.com/hacker.martin/", _
 &O0, &O0, 0)
 Unload Me
End Sub

Private Sub cmdyoutube_Click()
Dim X As String
 X = ShellExecute(Me.hwnd, "Open" _
 , "https://www.youtube.com/channel/UCEL746zBrw1bJMMkyDxgQAQ", _
 &O0, &O0, 0)
 Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = "Acerca de " & App.Title
    'lblVersion.Caption = "Versi�n " & App.Major & "." & App.Minor & "." & App.Revision
    cmdCambiar.ToolTipText = lenguaje_Menu(278)
    cargar_datos1
    Call cargarIdioma
    Me.Icon = frmprograma.Icon
    
    'carga Skins con el recurso del formulario requerido
    cargar_Skins Me
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener s�lo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versi�n conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox lenguaje_Menu(280), vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' C�digo de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tama�o de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tama�o de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, s�lo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversi�n...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor car�cter a car�cter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar despu�s de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vac�a
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function



Private Sub cargar_datos1()
Const espacio As String = "                               "
On Error GoTo no_se
ListView1.ListItems.Clear
    
      '  With ListView1.ColumnHeaders
      '      .Add , , "Recurso"
      '       .Add , , "Autores"
      '       End With
      ListView1.ColumnHeaders.Add , , lenguaje_Menu(308)
      ListView1.ColumnHeaders.Add , , lenguaje_Menu(309), 2700
 With ListView1
        ' Las pruebas ser�n en modo "detalle"
        .View = lvwReport
        ' al seleccionar un elemento, seleccionar la l�nea completa
        '.FullRowSelect = True
        ' Mostrar las l�neas de la cuadr�cula
       ' .GridLines = True
        ' No permitir la edici�n autom�tica del texto
        .LabelEdit = lvwManual
        ' Permitir m�ltiple selecci�n
        .MultiSelect = False
        ' Para que al perder el foco,
        ' se siga viendo el que est� seleccionado
        .HideSelection = False
   
             ListView1.View = lvwReport
             
     
                                      
          .ListItems.Add(, , lenguaje_Menu(283)).SubItems(1) = ":Martin Grasso Castillo."
                  .ListItems.Add(, , lenguaje_Menu(284)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(285)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(286)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(287)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(288)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(289)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(290)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(291)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(292)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(293)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(294)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(295)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(296)).SubItems(1) = " "" "
                  .ListItems.Add(, , lenguaje_Menu(297)).SubItems(1) = lenguaje_Menu(298)
                  .ListItems.Add(, , lenguaje_Menu(299)).SubItems(1) = lenguaje_Menu(300)
                  .ListItems.Add(, , lenguaje_Menu(301)).SubItems(1) = lenguaje_Menu(302)
                  .ListItems.Add(, , lenguaje_Menu(303)).SubItems(1) = lenguaje_Menu(304)
                  
     End With
      
no_se:
End Sub

Private Sub cargarIdioma()
  lblTitle.Caption = lenguaje_Menu(310)
  Lab1.Caption = lenguaje_Menu(282) & ": Canelones y en Tala,.*.* Club Atenas"
  lblDisclaimer.Caption = lenguaje_Menu(305)
  lblVersion.Caption = ""
  Me.Caption = lenguaje_Menu(281)
  cmdSysInfo.Caption = lenguaje_Menu(306)
  cmdOK.Caption = lenguaje_Menu(307)
End Sub
