VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Medicina del Trabajo"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18120
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   18120
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   9360
      Left            =   90
      TabIndex        =   2
      Top             =   945
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   16510
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSMonth SSMonth1 
         Height          =   2280
         Left            =   90
         TabIndex        =   3
         Top             =   6570
         Width           =   3345
         _Version        =   65537
         _ExtentX        =   5900
         _ExtentY        =   4022
         _StockProps     =   76
      End
      Begin VB.Image Image10 
         Height          =   375
         Left            =   90
         Picture         =   "Principal.frx":57E2
         Stretch         =   -1  'True
         Top             =   8865
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Libre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   4230
         Width           =   3210
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DEPURAR REGISTROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   270
         TabIndex        =   16
         Top             =   5760
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Image Image16 
         Height          =   780
         Left            =   1035
         Picture         =   "Principal.frx":79F1
         Stretch         =   -1  'True
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Libre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   3555
         Width           =   3210
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluació de Puesto de Trabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   645
         Left            =   180
         TabIndex        =   10
         Top             =   2070
         Width           =   3210
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trámite de Invalidéz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   2880
         Width           =   3210
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rev.27072021"
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   8865
         Width           =   1290
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   6120
         Width           =   3345
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Control Periódico de Salud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   1485
         Width           =   3210
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Programación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   225
         TabIndex        =   5
         Top             =   855
         Width           =   3165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1305
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   90
         Picture         =   "Principal.frx":9F5E
         Stretch         =   -1  'True
         Top             =   45
         Width           =   3390
      End
      Begin VB.Image Image4 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":3AB6F
         Stretch         =   -1  'True
         Top             =   630
         Width           =   3390
      End
      Begin VB.Image Image5 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":3B200
         Stretch         =   -1  'True
         Top             =   1305
         Width           =   3390
      End
      Begin VB.Image Image6 
         Height          =   555
         Left            =   90
         Picture         =   "Principal.frx":3B891
         Stretch         =   -1  'True
         Top             =   5985
         Width           =   3390
      End
      Begin VB.Image Image8 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":6C4A2
         Stretch         =   -1  'True
         Top             =   2655
         Width           =   3390
      End
      Begin VB.Image Image2 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":6CB33
         Stretch         =   -1  'True
         Top             =   1980
         Width           =   3390
      End
      Begin VB.Image Image17 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":6D1C4
         Stretch         =   -1  'True
         Top             =   3330
         Width           =   3390
      End
      Begin VB.Image Image9 
         Height          =   690
         Left            =   90
         Picture         =   "Principal.frx":6D855
         Stretch         =   -1  'True
         Top             =   4005
         Width           =   3390
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   3300
      Left            =   6255
      TabIndex        =   12
      Top             =   4185
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   5821
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Estadístico Por Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   315
         TabIndex        =   19
         Top             =   2025
         Width           =   2040
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Estadístico Por Médico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   315
         TabIndex        =   18
         Top             =   1395
         Width           =   2040
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Listado para afiliaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   465
         Left            =   315
         TabIndex        =   15
         Top             =   135
         Width           =   2040
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Estadístico Mensual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   315
         TabIndex        =   14
         Top             =   765
         Width           =   2040
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Revisión Médica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   270
         TabIndex        =   13
         Top             =   2655
         Width           =   2040
      End
      Begin VB.Image Image3 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6DEE6
         Stretch         =   -1  'True
         Top             =   45
         Width           =   2490
      End
      Begin VB.Image Image14 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6E577
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2490
      End
      Begin VB.Image Image15 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6EC08
         Stretch         =   -1  'True
         Top             =   2565
         Width           =   2490
      End
      Begin VB.Image Image11 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6F299
         Stretch         =   -1  'True
         Top             =   1305
         Width           =   2490
      End
      Begin VB.Image Image18 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6F92A
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   2490
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4515
      Left            =   3645
      TabIndex        =   20
      Top             =   1620
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   7964
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seguimiento de Trámite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   555
         Left            =   315
         TabIndex        =   27
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Image Image26 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":6FFBB
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   2490
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Direcciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   315
         TabIndex        =   26
         Top             =   4005
         Width           =   2040
      End
      Begin VB.Image Image25 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":7064C
         Stretch         =   -1  'True
         Top             =   3825
         Width           =   2490
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   315
         TabIndex        =   25
         Top             =   3375
         Width           =   2040
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   315
         TabIndex        =   24
         Top             =   2745
         Width           =   2040
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Asignación de turnos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   555
         Left            =   315
         TabIndex        =   23
         Top             =   1350
         Width           =   2040
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registro de exámenes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   315
         TabIndex        =   22
         Top             =   765
         Width           =   2040
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Programación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   315
         TabIndex        =   21
         Top             =   225
         Width           =   2040
      End
      Begin VB.Image Image19 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":70CDD
         Stretch         =   -1  'True
         Top             =   45
         Width           =   2490
      End
      Begin VB.Image Image20 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":7136E
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2490
      End
      Begin VB.Image Image23 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":719FF
         Stretch         =   -1  'True
         Top             =   1305
         Width           =   2490
      End
      Begin VB.Image Image24 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":72090
         Stretch         =   -1  'True
         Top             =   2565
         Width           =   2490
      End
      Begin VB.Image Image22 
         Height          =   645
         Left            =   45
         Picture         =   "Principal.frx":72721
         Stretch         =   -1  'True
         Top             =   3195
         Width           =   2490
      End
   End
   Begin VB.Image Image13 
      Height          =   870
      Left            =   13950
      Picture         =   "Principal.frx":72DB2
      Stretch         =   -1  'True
      Top             =   9405
      Width           =   4155
   End
   Begin VB.Image Image12 
      Height          =   5010
      Left            =   7830
      Picture         =   "Principal.frx":8069C
      Stretch         =   -1  'True
      Top             =   3645
      Width           =   5280
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Inactivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.Image Image21 
      Height          =   645
      Left            =   405
      Picture         =   "Principal.frx":81F04
      Stretch         =   -1  'True
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINA DEL TRABAJO PROGRAMACIÓN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   14355
      TabIndex        =   0
      Top             =   135
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image7 
      Height          =   870
      Left            =   0
      Picture         =   "Principal.frx":83F9E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18150
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Emp_Click()
If uactivado = 1 Then
    Load empresa
    empresa.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub equ_Click()
If uactivado = 1 Then
    Load Equipos
    Equipos.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Form_Activate()
If Label20.Caption = "Usuario Inactivo" Then
    Load Ingreso
    Ingreso.Show
End If
End Sub

Private Sub Form_Load()
'depurar

'DoEvents

'If siexiste = 1 Then
'    Image16.Visible = True
'    MsgBox "Existen registros para depurar", vbInformation, empresa
'    Label15.Visible = True
'Else
'    Image16.Visible = False
'End If
End Sub

Private Sub Image10_Click()
If uactivado = 1 Then
    Principal.Enabled = False
    Load Configuracion
    Configuracion.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If

End Sub

Private Sub Image16_Click()
If uactivado = 1 Then
    Principal.Enabled = False
    Load Depurados
    Depurados.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Image2_Click()
Dim Cn As New ADODB.Connection
Dim rsadm As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
rsadm.CursorType = adOpenKeyset
rsadm.LockType = adLockOptimistic
rsadm.ActiveConnection = Cn
rsadm.Source = "Select * from usuarios"
rsadm.Open

If Not rsadm.EOF Then
    Load Ingreso
    Ingreso.Show
Else
    origen = "verifadm"
    Principal.Enabled = False
    Load Usuarios
    Usuarios.Show
End If
End Sub
Private Sub Image3_Click()
uactivado = 0
Image2.Visible = True
Image3.Visible = False
Load Cerrado
Cerrado.Label2.Caption = vUsuario
Cerrado.Show
vUsuario = ""
vusucod = 0
Label22.Caption = "Usuario: Inactivo"
'Unload Llamar
'Set Llamar = Nothing
End Sub

Private Sub Label1_Click()
If uactivado = 1 Then
    MsgBox "En Implementación", vbInformation, empresa
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Label10_Click()
If uactivado = 1 Then
    Principal.Enabled = False
    Load AteMed
    AteMed.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If

End Sub

Private Sub Label11_Click()
Load EstMen
EstMen.Show
End Sub

Private Sub Label12_Click()
If uactivado = 1 Then
    MsgBox "En Implementación", vbInformation, empresa
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub
Private Sub Label13_Click()
If uactivado = 1 Then
    Principal.Enabled = False
    Load Arqueo
    Arqueo.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If

End Sub

Private Sub Label14_Click()
Load Programacion
Programacion.Show
End Sub

Private Sub Label16_Click()
If uactivado = 1 Then
    MsgBox "En Implementación", vbInformation, empresa
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Label17_Click()
Load Asignacion
Asignacion.Show
End Sub

Private Sub Label18_Click()
If vusuacc = "ADM" Then
'    SSFrame2.Visible = False
    Load NuevosIng
    NuevosIng.Show
Else
    MsgBox "Usuario sin acceso a esta opción", vbInformation, empresa
End If
End Sub

Private Sub Label19_Click()
Unload Principal
Set Principal = Nothing
End Sub

Private Sub Label20_Click()
Principal.Enabled = False
Load Ingreso
Ingreso.Show
End Sub

Private Sub Label21_Click()
    'SSFrame2.Visible = False
    SSFrame3.Visible = True
End Sub

Private Sub Label22_Click()
If vusuacc = "ADM" Then
    Load Empresas
    Empresas.Show
Else
    MsgBox "Usuario sin acceso a esta opción", vbInformation, empresa
End If
End Sub

Private Sub Label23_Click()
If vusuacc = "ADM" Then
    Load Direcciones
    Direcciones.Show
Else
    MsgBox "Usuario sin acceso a esta opción", vbInformation, empresa
End If
End Sub

Private Sub Label4_Click()
If uactivado = 1 Then
    SSFrame2.Visible = True
    SSFrame3.Visible = False
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Label5_Click()
If uactivado = 1 Then
    MsgBox "En Implementación", vbInformation, empresa
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Label7_Click()
If uactivado = 1 Then
    MsgBox "En Implementación", vbInformation, empresa
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Label8_Click()
Load EsteMed
EsteMed.Show
End Sub

Private Sub Label9_Click()
Load EstEmp
EstEmp.Show
End Sub

Private Sub Per_Click()
If uactivado = 1 Then
    Load pERSONAL
    pERSONAL.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub Repo_Click()

End Sub

Private Sub Sal_Click()
Unload Principal
Set Principal = Nothing
End Sub

Private Sub tip_Click()
If uactivado = 1 Then
    Load Tipos
    Tipos.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If

End Sub

Private Sub Trans_Click()

End Sub

Private Sub Uni_Click()
If uactivado = 1 Then
    Load Unidades
    Unidades.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If
End Sub

Private Sub usu_Click()
If uactivado = 1 Then
    Load Usuarios
    Usuarios.Show
Else
    MsgBox "No existe usuario activo", vbInformation, empresa
End If

End Sub

Private Sub Timer1_Timer()
Label14.Caption = Time$
End Sub
