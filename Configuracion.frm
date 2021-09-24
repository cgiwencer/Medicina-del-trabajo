VERSION 5.00
Begin VB.Form Configuracion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   2040
      Width           =   5580
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   5580
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   945
      TabIndex        =   6
      Top             =   2925
      Width           =   2025
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   3825
      Picture         =   "Configuracion.frx":0000
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3915
      TabIndex        =   5
      Top             =   2925
      Width           =   1710
   End
   Begin VB.Image Image11 
      Height          =   330
      Left            =   945
      Picture         =   "Configuracion.frx":2777
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección Medicina del Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   1800
      Width           =   3435
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Administrador Departamental"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1080
      Width           =   3435
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIGURACIÓN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   5130
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Configuracion.frx":44F2
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Configuracion.frx":EF4F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
   Begin VB.Image Image12 
      Height          =   555
      Left            =   765
      Picture         =   "Configuracion.frx":2C0B9
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   2790
      Width           =   2175
   End
   Begin VB.Image Image8 
      Height          =   465
      Left            =   3690
      Picture         =   "Configuracion.frx":2EAB3
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   2835
      Width           =   1905
   End
End
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Cn As New ADODB.Connection
Dim rsdi As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsdi.CursorType = adOpenKeyset
rsdi.LockType = adLockOptimistic
rsdi.ActiveConnection = Cn
rsdi.Source = "Select * from configuracion"
rsdi.Open

If Not rsdi.EOF Then
    Text1.Text = rsdi!ConfAdm
    Text2.Text = rsdi!ConfDirMD
End If
Cn.Close
End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload Configuracion
Set Configuracion = Nothing
End Sub

Private Sub Label2_Click()
If Len(Trim(Text1.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vConfAdm = Text1.Text
    vConfDirMD = Text2.Text
    
    Graba = "UPDATE configuracion SET ConfAdm = " & "'" & vConfAdm & "', ConfDirMD = " & "'" & vConfDirMD & "'"
    Cn.Execute Graba
    
    MsgBox "Se guardó la información", vbInformation, empresa
    Label10_Click
Else
    MsgBox "Debe ingresar el nombre del Administrador Departamental", vbInformation, empresa
    Text1.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HC0FFFF
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HC0FFFF
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
End Sub


