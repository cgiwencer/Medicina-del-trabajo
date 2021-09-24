VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Empresas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7890
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3390
      Top             =   3180
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=MedicinaT"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "MedicinaT"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "Select * from empresas ORDER BY empdes"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Empresa.frx":0000
      Height          =   4800
      Left            =   210
      TabIndex        =   1
      Top             =   810
      Width           =   8700
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      BevelColorFace  =   14737632
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   12632256
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns(0).Width=   14261
      Columns(0).Caption=   "Empresa"
      Columns(0).Name =   "empdes"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "empdes"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   15346
      _ExtentY        =   8467
      _StockProps     =   79
      Caption         =   " "
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   1296
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   5670
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
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
         Left            =   6300
         TabIndex        =   4
         Top             =   240
         Width           =   1785
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   6270
         Picture         =   "Empresa.frx":0015
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Image Image17 
         Height          =   465
         Left            =   6135
         Picture         =   "Empresa.frx":F87E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   135
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1095
      Left            =   9060
      TabIndex        =   5
      Top             =   6660
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1931
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Image Image5 
         Height          =   285
         Left            =   360
         Picture         =   "Empresa.frx":122B6
         Stretch         =   -1  'True
         Top             =   360
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
         Height          =   240
         Left            =   345
         TabIndex        =   6
         Top             =   375
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   225
         Picture         =   "Empresa.frx":14A2D
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   270
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1185
      Left            =   240
      TabIndex        =   7
      Top             =   6630
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   2090
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   5580
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   6885
         TabIndex        =   9
         Top             =   330
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   196608
         ForeColor       =   16777215
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Guardar y Corregir"
         ButtonStyle     =   4
         BevelWidth      =   0
      End
      Begin VB.Image Image11 
         Height          =   285
         Left            =   6390
         Picture         =   "Empresa.frx":1716E
         Stretch         =   -1  'True
         Top             =   375
         Width           =   285
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   270
         TabIndex        =   10
         Top             =   180
         Width           =   1950
      End
      Begin VB.Image Image12 
         Height          =   465
         Left            =   6135
         Picture         =   "Empresa.frx":18EE9
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   285
         Width           =   2535
      End
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Empresa.frx":1B8E3
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESAS"
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
      Height          =   285
      Left            =   9840
      TabIndex        =   0
      Top             =   210
      Width           =   1365
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Empresa.frx":26340
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "Empresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vEmpDes As String
Public vEmpId As Integer
Private Sub Label10_Click()
Unload Empresas
Set Empresas = Nothing
End Sub

Private Sub Label24_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = UCase(Text8.Text)
    Adodc1.RecordSource = "SELECT * from empresas WHERE  EmpDes LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
Else
    Adodc1.RecordSource = "SELECT * from empresas ORDER BY EmpDes"
    Adodc1.Refresh
End If

End Sub

Private Sub SSCommand1_Click()
If seleccion = 1 Then
    Dim Cn As New ADODB.Connection
    Dim rsem As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vEmpDes = Text1.Text
    
    rsem.CursorType = adOpenKeyset
    rsem.LockType = adLockOptimistic
    rsem.ActiveConnection = Cn
    rsem.Source = "Select * from empresas where EmpDes = " & "'" & vEmpDes & "'"
    rsem.Open

    If rsem.EOF Then
        actualizaempresa
    Else
        If MsgBox("Empresa ya existente. Desea cambiar la razón Social de todas formas?", vbYesNo, empresa) = vbYes Then
            actualizaempresa
            borra = "DELETE FROM empresas WHERE EmpId = " & vEmpId
            Cn.Execute borra
            Adodc1.Refresh
        End If
        End If
Else
    MsgBox "Seleccione una empresa de la lista"
End If
End Sub
Private Sub SSOleDBGrid1_Click()
seleccion = 1
Text1.Text = Adodc1.Recordset.Fields("EmpDes")
vempresades = Text1.Text
vEmpId = Adodc1.Recordset.Fields("EmpId")
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = &HC0FFFF
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text8_GotFocus()
Text8.BackColor = &HC0FFFF
End Sub
Private Sub Text8_LostFocus()
Text8.BackColor = &HFFFFFF
Text8.Text = UCase(Text8.Text)
End Sub

Private Function actualizaempresa()
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
'Graba cambio
cambia = "UPDATE empresas set EmpDes = " & "'" & vEmpDes & "' WHERE EmpDes = " & "'" & vempresades & "'"
Cn.Execute cambia
'Cambia las coincidencias
corrige = "UPDATE NUEVOS set EmpDes = " & "'" & vEmpDes & "' WHERE EmpDes = " & "'" & vempresades & "'"
Cn.Execute corrige
Cn.Close
Adodc1.Refresh
MsgBox "Se cambio el nombre y se corrigieron coincidencias"
End Function
