VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Depurados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7395
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3150
      Top             =   3555
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
      RecordSource    =   "Select * from depurados ORDER BY NueNom"
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
      Bindings        =   "Depurados.frx":0000
      Height          =   4305
      Left            =   90
      TabIndex        =   0
      Top             =   855
      Width           =   12015
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      BevelColorFace  =   8421376
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
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   7408
      Columns(0).Caption=   "Personas"
      Columns(0).Name =   "NueNom"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "NueNom"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7170
      Columns(1).Caption=   "Empresa"
      Columns(1).Name =   "EmpDes"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "EmpDes"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2831
      Columns(2).Caption=   "Fecha Examen"
      Columns(2).Name =   "NueFeM"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "NueFeM"
      Columns(2).DataType=   7
      Columns(2).NumberFormat=   "dd/mm/yyyy"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2752
      Columns(3).Caption=   "Fecha Depuración"
      Columns(3).Name =   "NueFeD"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "NueFeD"
      Columns(3).DataType=   7
      Columns(3).NumberFormat=   "dd/mm/yyyy"
      Columns(3).FieldLen=   256
      _ExtentX        =   21193
      _ExtentY        =   7594
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
      ForeColor       =   16777215
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
      Left            =   135
      TabIndex        =   4
      Top             =   5355
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1296
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   7080
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
         Left            =   7920
         TabIndex        =   6
         Top             =   270
         Width           =   1275
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   7470
         Picture         =   "Depurados.frx":0015
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Image Image17 
         Height          =   465
         Left            =   7335
         Picture         =   "Depurados.frx":F87E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   135
         Width           =   1905
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "     Depurar     Selección"
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
      Height          =   510
      Left            =   12510
      TabIndex        =   7
      Top             =   2700
      Width           =   2265
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   12555
      Picture         =   "Depurados.frx":122B6
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   12600
      Picture         =   "Depurados.frx":21AA2
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "    Salir"
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
      Left            =   12870
      TabIndex        =   3
      Top             =   4095
      Width           =   1410
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "     Depurar   Todos"
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
      Height          =   510
      Left            =   12600
      TabIndex        =   2
      Top             =   1530
      Width           =   1950
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   12600
      Picture         =   "Depurados.frx":24219
      Stretch         =   -1  'True
      Top             =   2790
      Width           =   330
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAS POR DEPURAR"
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
      Left            =   11655
      TabIndex        =   1
      Top             =   225
      Width           =   2820
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Depurados.frx":271DF
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   45
      Picture         =   "Depurados.frx":31C3C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14595
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   12465
      Picture         =   "Depurados.frx":4EDA6
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   1485
      Width           =   1905
   End
   Begin VB.Image Image8 
      Height          =   510
      Left            =   12465
      Picture         =   "Depurados.frx":5135B
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   3960
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   12465
      Picture         =   "Depurados.frx":53A9C
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   2655
      Width           =   1905
   End
End
Attribute VB_Name = "Depurados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " " & "personas"
End Sub

Private Sub Label1_Click()
If seleccion = 1 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    depura = "UPDATE nuevos SET NueEst = " & 3 & " WHERE NueId = " & vnueid
    Cn.Execute depura
    
    depura = "DELETE FROM depurados WHERE NueId = " & vnueid
    Cn.Execute depura
    
    MsgBox "Proceso concluido", vbInformation, empresa
    seleccion = 0
    Adodc1.Refresh
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload Depurados
Set Depurados = Nothing
End Sub
Private Sub Label24_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from depurados WHERE  NueNom LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
Else
    Adodc1.RecordSource = "SELECT * from depurados"
    Adodc1.Refresh
End If

End Sub
Private Sub Label8_Click()
If MsgBox("Esta seguro de depurar a esta(s) persona(s) ?", vbYesNo, empresa) = vbYes Then
    Dim Cn As New ADODB.Connection
    Dim rsd As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsd.CursorType = adOpenKeyset
    rsd.LockType = adLockOptimistic
    rsd.ActiveConnection = Cn
    rsd.Source = "Select * from depurados"
    rsd.Open
    
    If Not rsd.EOF Then
        Do While Not rsd.EOF
            vnueid = rsd!NueId
            depura = "UPDATE nuevos SET NueEst = " & 3 & " WHERE NueId = " & vnueid
            Cn.Execute depura
            rsd.MoveNext
        Loop
    End If
    MsgBox "Proceso concluido", vbInformation, empresa
    Label10_Click
    Principal.Enabled = True
    Principal.Image16.Visible = False
    Principal.Label15.Visible = False
Else
    Label10_Click
    Principal.Enabled = True
End If
End Sub

Private Sub SSOleDBGrid1_Click()
seleccion = 1
vnueid = Adodc1.Recordset.Fields("NueId")
End Sub

