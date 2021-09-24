VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Ingreso 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   4500
   ClientTop       =   4755
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1170
      TabIndex        =   4
      Text            =   "Nombre de usuario"
      Top             =   1260
      Width           =   3165
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1755
      Width           =   3165
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   420
      Left            =   225
      TabIndex        =   0
      Top             =   2700
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Ingresar"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   2070
      TabIndex        =   1
      Top             =   2700
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "C&ambiar clave"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   420
      Left            =   3915
      TabIndex        =   2
      Top             =   2700
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210816
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancelar"
      BevelWidth      =   1
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   720
      Picture         =   "Ingreso.frx":0000
      Stretch         =   -1  'True
      Top             =   1755
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   720
      Picture         =   "Ingreso.frx":0F16
      Stretch         =   -1  'True
      Top             =   1215
      Width           =   375
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Ingreso.frx":203C
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA PETROLERA DE SALUD - LA PAZ"
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
      Left            =   765
      TabIndex        =   3
      Top             =   225
      Width           =   4200
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Ingreso.frx":CA99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5505
   End
End
Attribute VB_Name = "Ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acceso, vfoto As String
Private Sub Form_Load()
KeyPreview = True
End Sub
Sub ChgEnterToTab(KeyCode As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   SendKeys "{TAB}"
   KeyAscii = 0
End If
End Sub
Private Function accesosautorizados()
Dim Cn As New ADODB.Connection
Dim rsd As New ADODB.Recordset
Dim rsp As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
'Seleccion de tabla Derechos
rsd.CursorType = adOpenKeyset
rsd.LockType = adLockOptimistic
rsd.ActiveConnection = Cn
rsd.Source = "Select * from derechos"
rsd.Open
'Seleccion de tabla Procesos
rsp.CursorType = adOpenKeyset
rsp.LockType = adLockOptimistic
rsp.ActiveConnection = Cn
rsp.Source = "Select * from procesos"
rsp.Open

intento = 0

If Not rsd.EOF Then
    rsd.MoveFirst
    Do While Not rsd.EOF
        If rsd!UsuCod = vusucod Then
            n = 0
            Do
                c = "P"
                n = n + 1
                d = Trim(c) + Trim(Str(n))
                If rsd.Fields(d) > 0 Then
                    rsp.MoveFirst
                    Do While Not rsp.EOF
                        If rsp!PrcNum = rsd.Fields(d) Then
                            MsgBox "Se habilita el Proceso: " & rsp!PrcDes
                            siaccesos = 0
                            Exit Do
                        End If
                        rsp.MoveNext
                    Loop
                End If
            Loop Until rsd.Fields(d) = 0
            Exit Do
        Else
            siaccesos = 1
        End If
        rsd.MoveNext
    Loop
End If
If siaccesos = 0 Then
    'Boton3_Click
Else
    MsgBox "Usuario sin derechos asignados", vbInformation, "Consulte con un supervisor"
    'Boton3_Click
End If
End Function
Private Function verificadatos()
'**** Conección a la Tabla Usuarios ***
Dim Cn As New ADODB.Connection
Dim rsu As New ADODB.Recordset
Cn.ConnectionString = Cadena

Cn.Open
rsu.CursorType = adOpenKeyset
rsu.LockType = adLockOptimistic
rsu.ActiveConnection = Cn
rsu.Source = "Select * from usuarios"
rsu.Open

'**** Inicio verificación datos ****
If Not rsu.EOF > 0 Then
    rsu.MoveFirst
    'Do While intento < 3
        Do While Not rsu.EOF
            If rsu!Usu_Usu = Text1.Text Then
                siexiste = 0
                If rsu!Usu_Cla = Trim(Text2.Text) Then
                    siexiste = 0
                    vUsuario = rsu!Usu_Nom
                    vusucod = rsu!Usu_Id
'                    vusuhor = rsu!Usu_hor
                    vusuacc = rsu!Usu_Acc
                    Exit Do
                Else
                    siexiste = 1
                    Exit Do
                End If
            Else
                siexiste = 1
            End If
            rsu.MoveNext
        Loop
        'If siexiste = 1 Then
        '    Exit Do
        'Else
        '    Exit Do
        'End If
    'Loop
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
Principal.Enabled = True
Unload Ingreso
Set Ingreso = Nothing
End Sub
Private Sub SSCommand1_Click()
verificadatos
If siexiste = 1 Then
    MsgBox "Datos incorrectos", vbInformation, "Por favor revise"
ElseIf siexiste = 0 Then
    Ingreso.Enabled = False
    Load Cambiarclave
    Cambiarclave.Show
End If

End Sub

Private Sub SSCommand2_Click()
Principal.Enabled = True
Unload Ingreso
Set Ingreso = Nothing
Unload Principal
Set Principal = Nothing
End Sub

Private Sub SSCommand7_Click()
verificadatos
If siexiste = 0 Then
    intento = 0
End If
If intento = 1 Then
    MsgBox "Si usted olvido su clave, por favor consulte con el administrador del sistema, caso contrario su clave será bloqueada", vbCritical, "Ultimo oportunidad"
ElseIf intento = 3 Then
    MsgBox "Consulte con el administrador del sistema", vbCritical, "Clave bloqueada"
End If

If siexiste = 1 Then

    Load NoAutorizado
'    NoAutorizado.Label1.Caption = Label4.Caption
    NoAutorizado.Show

    ''ghistorico
    vCambiar = 1
    'MsgBox "Por favor revise", vbInformation, "Datos incorrectos"
    intento = intento + 1
    
    If intento = 3 Then
        MsgBox "Consulte con el supervisor", vbInformation, "Llegó al límite de intentos"
        Menup.Enabled = True
        Unload Ingreso
        Set Ingreso = Nothing
    End If
ElseIf siexiste = 0 Then
    ''ghistorico
    vCambiar = 0
'    accesosautorizados
    If siaccesos = 1 Then
     'No tiene derechos
    ElseIf siaccesos = 0 Then
     'Si tiene derechos
        Load autorizado
        autorizado.Show
        
        uactivado = 1
        'MenuP.Label1.Caption = ""
        'Menup.Image3.Visible = False
        'Menup.Label4.Visible = False
        'Menup.Image4.Left = 225
        'Menup.Label5.Left = 675
        'Menup.Image4.Visible = True
        'Menup.Label5.Visible = True
        Principal.Label20.Caption = vUsuario
        'If Len(Trim(vfoto)) > 0 Then
        '    Menup.Image11.Picture = LoadPicture(App.Path & "\fotosusuarios\" & vfoto)
        'Else
         '   Menup.Image11.Picture = LoadPicture(App.Path & "\fotosusuarios\siusuario.jpg")
        'End If
    End If
End If
Unload Me
End Sub
Private Sub SSOleDBGrid1_DblClick()
Label4.Caption = SSOleDBGrid1.Columns(0).Value
Text2.SetFocus
End Sub
Private Function ghistorico()
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("HisFec") = Date
If siexiste = 1 Then
    Adodc2.Recordset.Fields("HisDes") = "El usuario " & Label4.Caption & " intento ingresar al sistema"
ElseIf siexiste = 0 Then
    Adodc2.Recordset.Fields("HisDes") = "El usuario " & Label4.Caption & " ingreso al sistema"
End If
If Mid(Time, 2, 1) <> ":" Then
    hora = Left(Time, 2) & ":" & Mid(Time, 4, 2)
Else
    hora = Left(Time, 1) & ":" & Mid(Time, 3, 2)
End If
Adodc2.Recordset.Fields("HisHor") = hora
Adodc2.Recordset.Update
End Function


Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub
