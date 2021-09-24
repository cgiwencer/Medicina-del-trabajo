VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form EstEmp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   900
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2220
      Left            =   225
      TabIndex        =   4
      Top             =   855
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   3916
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   375
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "EstEmp.frx":0000
         Left            =   4095
         List            =   "EstEmp.frx":0013
         TabIndex        =   3
         Top             =   1125
         Width           =   1500
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   2145
         TabIndex        =   2
         Top             =   1110
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   270
         TabIndex        =   1
         Top             =   1110
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   90
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   3285
         Picture         =   "EstEmp.frx":0035
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   630
         TabIndex        =   9
         Top             =   1725
         Width           =   1815
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   540
         Picture         =   "EstEmp.frx":27AC
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Gestión"
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
         Height          =   195
         Left            =   3975
         TabIndex        =   8
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio"
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
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
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
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   885
         Width           =   1005
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
         Left            =   3330
         TabIndex        =   5
         Top             =   1710
         Width           =   1920
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   405
         Picture         =   "EstEmp.frx":580E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1620
         Width           =   1965
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   3150
         Picture         =   "EstEmp.frx":7DC3
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1620
         Width           =   1965
      End
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "EstEmp.frx":A504
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTE POR EMPRESA"
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
      Left            =   3465
      TabIndex        =   10
      Top             =   225
      Width           =   2820
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "EstEmp.frx":14F61
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "EstEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_GotFocus()
Dim lResp As Long
lResp = SendMessageLong(Combo1.hWnd, &H14F, True, 0)

If Len(Trim(Combo1.Text)) > 0 Then
    vempresa = Combo1.Text
End If

Dim Cn As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsp.CursorType = adOpenKeyset
rsp.LockType = adLockOptimistic
rsp.ActiveConnection = Cn
rsp.Source = "Select * from empresas"
rsp.Open

If Not rsp.EOF Then
    rsp.MoveFirst
    Combo1.Clear
    Do While Not rsp.EOF
        Combo1.AddItem UCase(rsp!EmpDes)
        rsp.MoveNext
    Loop
End If
If Len(Trim(vempresa)) > 0 Then
    Combo1.Text = vempresa
End If
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1.Text)
Cn.Close

End Sub

Private Sub Form_Load()
Combo2.ListIndex = 0
End Sub

Private Sub Label1_Click()
Dim total As Integer
vEmpDes = Combo1.Text
If Len(Trim(Combo2.Text)) > 0 Then
    vfechai = SSDateCombo1.Date
    vfechaf = SSDateCombo2.Date
    vfechai = Format(vfechai, "YYYY-MM-dd")
    vfechaf = Format(vfechaf, "yyyy-mm-dd")
    vfechair = SSDateCombo1.Date
    vfechafr = SSDateCombo2.Date
    vges = Combo2.Text
         
    Dim Cn As New ADODB.Connection
    Dim rsla As New ADODB.Recordset ' Recordset de reporte estadistico
    Cn.ConnectionString = Cadena
    Cn.Open
    
    borra = "DELETE FROM imprepemp"
    Cn.Execute borra
    
    Graba = "INSERT INTO imprepemp Select * from nuevos where NueFeS BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(FecRev) = " & vges & " AND EmpDes = " & "'" & vEmpDes & "'"
    Cn.Execute Graba
    
    CrystalReport1.ReportFileName = App.Path & "\PorEmpresa.rpt"
    CrystalReport1.Formulas(1) = "del = " & "'" & vfechair & "'"
    CrystalReport1.Formulas(0) = "al = " & "'" & vfechafr & "'"
    CrystalReport1.Action = 1
Else
    MsgBox "Debe seleccionar una gestión", vbInformation, empresa
End If

End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload EstEmp
Set EstEmp = Nothing
End Sub
