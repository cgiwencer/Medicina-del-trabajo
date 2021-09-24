VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form AteMed 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4410
      Top             =   765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1365
      Left            =   180
      TabIndex        =   0
      Top             =   930
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   2408
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   435
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         TabIndex        =   4
         Top             =   210
         Width           =   540
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   2250
         Picture         =   "AteMed.frx":0000
         Stretch         =   -1  'True
         Top             =   255
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
         Left            =   2310
         TabIndex        =   3
         Top             =   255
         Width           =   1815
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
         Left            =   2295
         TabIndex        =   2
         Top             =   735
         Width           =   1920
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   2250
         Picture         =   "AteMed.frx":3062
         Stretch         =   -1  'True
         Top             =   735
         Width           =   345
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   2115
         Picture         =   "AteMed.frx":57D9
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   165
         Width           =   1965
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   2115
         Picture         =   "AteMed.frx":7D8E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   645
         Width           =   1965
      End
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "AteMed.frx":A4CF
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE ATENCIÒN MÈDICA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1710
      TabIndex        =   5
      Top             =   225
      Width           =   3450
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "AteMed.frx":14F2C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4920
   End
End
Attribute VB_Name = "AteMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Dim Cn As New ADODB.Connection
Dim rsar As New ADODB.Recordset ' Recordset de atemed
Cn.ConnectionString = Cadena
Cn.Open
vfechai = SSDateCombo1.Date
vfechai = Format(vfechai, "YYYY-MM-dd")

borrat = "DELETE FROM LisAteMed"
Cn.Execute borrat

grabalista = "INSERT INTO LisAteMed (NueNom, TelPac, TelEnc, EmpDes) select NueNom, NueTeI, NueTer, EmpDes from nuevos where NueFeM = " & "'" & vfechai & "'"
Cn.Execute grabalista

vfechaini = SSDateCombo1.Text

CrystalReport1.ReportFileName = App.Path & "\listadoAM.rpt"
CrystalReport1.Formulas(0) = "fecha = " & "'" & vfechaini & "'"
CrystalReport1.Action = 1

End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload AteMed
Set AteMed = Nothing
End Sub
