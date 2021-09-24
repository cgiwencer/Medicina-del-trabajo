VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form EsteMed 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6390
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
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1950
      Left            =   225
      TabIndex        =   3
      Top             =   855
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   3440
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "EsteMed.frx":0000
         Left            =   4050
         List            =   "EsteMed.frx":0013
         TabIndex        =   2
         Top             =   630
         Width           =   1500
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   2100
         TabIndex        =   1
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   225
         TabIndex        =   0
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
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
         TabIndex        =   8
         Top             =   1350
         Width           =   1920
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
         Left            =   2340
         TabIndex        =   7
         Top             =   390
         Width           =   1005
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
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   1335
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
         Left            =   3930
         TabIndex        =   5
         Top             =   405
         Width           =   1665
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   540
         Picture         =   "EsteMed.frx":0035
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   285
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
         TabIndex        =   4
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   3285
         Picture         =   "EsteMed.frx":3097
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   345
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   3150
         Picture         =   "EsteMed.frx":580E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1260
         Width           =   1965
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   405
         Picture         =   "EsteMed.frx":7F4F
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1260
         Width           =   1965
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTE ESTADISTICO POR MEDICO"
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
      Left            =   2295
      TabIndex        =   9
      Top             =   225
      Width           =   4125
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "EsteMed.frx":A504
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "EsteMed.frx":14F61
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "EsteMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo2.ListIndex = 0
End Sub

Private Sub Label1_Click()
Dim total As Integer
total = 0
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
    
    borra = "DELETE FROM imprepmed"
    Cn.Execute borra
    
    graba = "INSERT INTO imprepmed Select * from nuevos where FecRev BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(FecRev) = " & vges & " AND NueFic > " & 0
    Cn.Execute graba
    
    CrystalReport1.ReportFileName = App.Path & "\PorMedico.rpt"
    CrystalReport1.Formulas(1) = "del = " & "'" & vfechair & "'"
    CrystalReport1.Formulas(0) = "al = " & "'" & vfechafr & "'"
    CrystalReport1.Action = 1
Else
    MsgBox "Debe seleccionar una gestión", vbInformation, empresa
End If
End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload EsteMed
Set EsteMed = Nothing
End Sub
